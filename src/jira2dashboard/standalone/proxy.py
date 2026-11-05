"""The proxy of the Windows settings for the Jira site, as the browser finds it: what the VBA module modProxy does.

The automatic detection (WPAD) and the proxy script (PAC, AutoConfigURL) through WinHttpGetProxyForUrl, then the fixed
proxy of the user's settings (WinHttpGetIEProxyConfigForCurrentUser), else direct. Strings that WinHTTP allocates
are copied, then freed with GlobalFree. The structures are ctypes structures (64 and 32-bit).
"""
import ctypes
import sys
from ctypes import wintypes
from typing import Callable, NamedTuple

from jira2dashboard.standalone.jira_api import normalize_proxy, trim_white

AUTOPROXY_AUTO_DETECT, AUTOPROXY_CONFIG_URL = 1, 2
AUTO_DETECT_TYPE_DHCP, AUTO_DETECT_TYPE_DNS_A = 1, 2
ACCESS_TYPE_NO_PROXY = 1


class Settings(NamedTuple):
    """The proxy settings of the user in Windows (Internet Options)."""
    auto_detect: bool
    auto_config_url: str
    proxy: str
    bypass: str


class Script(NamedTuple):
    ok: bool
    proxy: str
    bypass: str
    error: int


class Detected(NamedTuple):
    server: str                 # "" = direct
    bypass: str
    mode: str                   # how it was found, for the log


class _IEProxyConfig(ctypes.Structure):
    _fields_ = [("fAutoDetect", wintypes.BOOL), ("lpszAutoConfigUrl", ctypes.c_void_p),
                ("lpszProxy", ctypes.c_void_p), ("lpszProxyBypass", ctypes.c_void_p)]


class _AutoProxyOptions(ctypes.Structure):
    _fields_ = [("dwFlags", wintypes.DWORD), ("dwAutoDetectFlags", wintypes.DWORD),
                ("lpszAutoConfigUrl", wintypes.LPCWSTR), ("lpvReserved", ctypes.c_void_p),
                ("dwReserved", wintypes.DWORD), ("fAutoLogonIfChallenged", wintypes.BOOL)]


class _ProxyInfo(ctypes.Structure):
    _fields_ = [("dwAccessType", wintypes.DWORD), ("lpszProxy", ctypes.c_void_p), ("lpszProxyBypass", ctypes.c_void_p)]


def _api():
    """winhttp and kernel32 with their argument types (64-bit pointers must not be cut to 32 bits)."""
    winhttp, kernel32 = ctypes.WinDLL("winhttp", use_last_error=True), ctypes.WinDLL("kernel32")
    kernel32.GlobalFree.argtypes, kernel32.GlobalFree.restype = [ctypes.c_void_p], ctypes.c_void_p
    winhttp.WinHttpGetIEProxyConfigForCurrentUser.argtypes = [ctypes.POINTER(_IEProxyConfig)]
    winhttp.WinHttpGetIEProxyConfigForCurrentUser.restype = wintypes.BOOL
    winhttp.WinHttpOpen.argtypes = [wintypes.LPCWSTR, wintypes.DWORD, wintypes.LPCWSTR, wintypes.LPCWSTR,
                                    wintypes.DWORD]
    winhttp.WinHttpOpen.restype = ctypes.c_void_p
    winhttp.WinHttpGetProxyForUrl.argtypes = [ctypes.c_void_p, wintypes.LPCWSTR, ctypes.POINTER(_AutoProxyOptions),
                                              ctypes.POINTER(_ProxyInfo)]
    winhttp.WinHttpGetProxyForUrl.restype = wintypes.BOOL
    winhttp.WinHttpCloseHandle.argtypes, winhttp.WinHttpCloseHandle.restype = [ctypes.c_void_p], wintypes.BOOL
    return winhttp, kernel32


def _take(kernel32, pointer) -> str:
    """A string that WinHTTP allocated: copied, then freed."""
    if not pointer:
        return ""
    text = ctypes.wstring_at(pointer)
    kernel32.GlobalFree(pointer)
    return text


def windows_settings() -> Settings:
    """The proxy settings of the user; all empty when Windows has none for this user."""
    if sys.platform != "win32":
        return Settings(False, "", "", "")
    winhttp, kernel32 = _api()
    config = _IEProxyConfig()
    if not winhttp.WinHttpGetIEProxyConfigForCurrentUser(ctypes.byref(config)):
        return Settings(False, "", "", "")
    return Settings(bool(config.fAutoDetect), _take(kernel32, config.lpszAutoConfigUrl),
                    _take(kernel32, config.lpszProxy), _take(kernel32, config.lpszProxyBypass))


def script_proxy(url: str, auto_detect: bool, auto_config_url: str) -> Script:
    """The proxy that WPAD and/or the PAC script give for `url` (one call with both flags when both are set)."""
    winhttp, kernel32 = _api()
    session = winhttp.WinHttpOpen("jira2dashboard", ACCESS_TYPE_NO_PROXY, None, None, 0)
    if not session:
        return Script(False, "", "", ctypes.get_last_error())
    try:
        options = _AutoProxyOptions()
        if auto_detect:
            options.dwFlags |= AUTOPROXY_AUTO_DETECT
            options.dwAutoDetectFlags = AUTO_DETECT_TYPE_DHCP | AUTO_DETECT_TYPE_DNS_A
        if auto_config_url:
            options.dwFlags |= AUTOPROXY_CONFIG_URL
            options.lpszAutoConfigUrl = auto_config_url
        options.fAutoLogonIfChallenged = True
        info = _ProxyInfo()
        if not winhttp.WinHttpGetProxyForUrl(session, url, ctypes.byref(options), ctypes.byref(info)):
            return Script(False, "", "", ctypes.get_last_error())
        return Script(True, _take(kernel32, info.lpszProxy), _take(kernel32, info.lpszProxyBypass), 0)
    finally:
        winhttp.WinHttpCloseHandle(session)


def detect_proxy(url: str, settings: Settings | None = None,
                 script: Callable[[str, bool, str], Script] = script_proxy) -> Detected:
    """The proxy for `url`: WPAD or PAC first, then the fixed proxy, else direct (modProxy.DetectProxy)."""
    ie = settings if settings is not None else windows_settings()
    how = ""
    if ie.auto_detect or ie.auto_config_url:
        how = "proxy script (PAC)" if ie.auto_config_url else "automatic detection (WPAD)"
        found = script(url, ie.auto_detect, ie.auto_config_url)
        if found.ok:
            return Detected(found.proxy, found.bypass, how + (": a proxy" if found.proxy else ": direct"))
        how += f" found nothing (Windows error {found.error}), then "
    if ie.proxy:
        return Detected(ie.proxy, ie.bypass, how + "the fixed proxy of the Windows settings")
    return Detected("", "", how + "direct (no proxy in the Windows settings)")


def choose_proxy(server: str, bypass: str, url: str,
                 detect: Callable[[str], Detected] = detect_proxy) -> Detected:
    """The proxy of a run from Config (modJira.JiraConfigFromSheet): DIRECT = none; a typed one; empty = the Windows
    settings, whose bypass list is used only when none is typed."""
    typed, typed_bypass = trim_white(server), trim_white(bypass)
    if typed.upper() == "DIRECT":
        return Detected("", typed_bypass, "DIRECT in Config: no proxy")
    if typed:
        return Detected(normalize_proxy(typed), typed_bypass, "typed in Config")
    found = detect(url)
    return Detected(normalize_proxy(found.server), typed_bypass or found.bypass, found.mode)
