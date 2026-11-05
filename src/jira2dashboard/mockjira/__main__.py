"""Run the fake Jira for the VBA client:  scripts\\run.ps1 mock-jira [options]

Put the printed BaseUrl, email and token in the Config sheet. Stop with Ctrl+C (or scripts\\stop-mock.ps1 when
scripts\\start-mock.ps1 started it).
The options mirror the parameters of the existing PowerShell extraction script (project keys, the ids of the
Business Value / Story Points / PI custom fields) and its credentials (Basic auth: email + API token).
"""
import argparse
import time

from jira2dashboard.mockjira.anomalies import with_anomalies
from jira2dashboard.mockjira.dataset import Dataset, FieldIds
from jira2dashboard.mockjira.generator import generate
from jira2dashboard.mockjira.server import MOCK_EMAIL, MOCK_TOKEN, MockJira
from jira2dashboard.params.model import project_keys


def parse(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--issues", type=int, default=1200)
    parser.add_argument("--seed", type=int, default=1)
    parser.add_argument("--port", type=int, default=8099)
    parser.add_argument("--projects", default=",".join(project_keys()),
                        help="project keys, comma separated; the FIRST one is the Program (features). "
                             "Default: params/projects.md")
    parser.add_argument("--bv-field", default=FieldIds.business_value, help="id of the Business Value field")
    parser.add_argument("--sp-field", default=FieldIds.story_points, help="id of the Story Points field")
    parser.add_argument("--pi-field", default=FieldIds.pi, help="id of the PI field")
    parser.add_argument("--anomalies", action="store_true",
                        help="add the known data-quality cases of the tests, so that the data checks find something")
    parser.add_argument("--email", default=MOCK_EMAIL)
    parser.add_argument("--token", default=MOCK_TOKEN)
    parser.add_argument("--no-auth", action="store_true", help="accept any credentials")
    parser.add_argument("--proxy-challenge", default=None, metavar="SCHEME",
                        help="act as a proxy that asks for authentication (Negotiate, NTLM, ...): 407 until "
                             "the request carries Proxy-Authorization")
    parser.add_argument("--page-clamp", type=int, default=100, help="most issues a search page can hold")
    return parser.parse_args(argv)


def dataset_for(args: argparse.Namespace) -> Dataset:
    fields = FieldIds(story_points=args.sp_field, business_value=args.bv_field, pi=args.pi_field)
    projects = tuple(p.strip() for p in args.projects.split(",") if p.strip())
    dataset = generate(seed=args.seed, issues=args.issues, projects=projects, fields=fields)
    return with_anomalies(dataset, projects)[0] if args.anomalies else dataset


def main(argv: list[str] | None = None) -> int:
    args = parse(argv)
    dataset = dataset_for(args)
    fields = dataset.fields
    credentials = None if args.no_auth else (args.email, args.token)
    with MockJira(dataset, port=args.port, credentials=credentials, page_clamp=args.page_clamp,
                  proxy_challenge=args.proxy_challenge) as jira:
        extra = " with the known anomalies" if args.anomalies else ""
        print(f"Fake Jira: {len(dataset.issues)} issues in {', '.join(dataset.projects())}{extra} (seed {args.seed}).")
        print(f"  BaseUrl : {jira.base_url}")
        print("  Auth    : " + ("none checked" if credentials is None else f"Basic, email {args.email}, token {args.token}"))
        print(f"  Fields  : story points {fields.story_points}, business value {fields.business_value}, PI {fields.pi}")
        print("Ctrl+C to stop.", flush=True)
        try:
            while True:
                time.sleep(1)
        except KeyboardInterrupt:
            print("stopped")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
