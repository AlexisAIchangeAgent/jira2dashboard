"""The oracle (L0) of the scenario art-metrics: every number worked out BY HAND from AL's text (2026-11-16), never computed by the code
that is tested (mockjira/art_metrics.py, flow/, the formulas). Unit tests, the Python-path scenario and the workbook scenario compare to it.
"""
import datetime as dt

# --- the calendar: [Start, End), PIs of 18 weeks, iterations of 3 weeks, no gap --------------------------------------------------------
PIS = {"PI-1": ("2025-01-01", "2025-05-07"), "PI-2": ("2025-05-07", "2025-09-10"),
       "PI-3": ("2025-09-10", "2026-01-14"), "PI-4": ("2026-01-14", "2026-05-20")}
ITERATION_STARTS = {
    "IT-1-1": "2025-01-01", "IT-1-2": "2025-01-22", "IT-1-3": "2025-02-12", "IT-1-4": "2025-03-05", "IT-1-5": "2025-03-26", "IT-1-6": "2025-04-16",
    "IT-2-1": "2025-05-07", "IT-2-2": "2025-05-28", "IT-2-3": "2025-06-18", "IT-2-4": "2025-07-09", "IT-2-5": "2025-07-30", "IT-2-6": "2025-08-20",
    "IT-3-1": "2025-09-10", "IT-3-2": "2025-10-01", "IT-3-3": "2025-10-22", "IT-3-4": "2025-11-12", "IT-3-5": "2025-12-03", "IT-3-6": "2025-12-24",
    "IT-4-1": "2026-01-14", "IT-4-2": "2026-02-04", "IT-4-3": "2026-02-25", "IT-4-4": "2026-03-18", "IT-4-5": "2026-04-08", "IT-4-6": "2026-04-29"}
AS_OF = dt.datetime(2026, 5, 20)

# --- where each Feature is released or deferred, its team ---------------------------------------------------------------------
RELEASED = {1: "IT-1-1", 2: "IT-1-2", 3: "IT-1-3", 4: "IT-1-4", 5: "IT-1-5", 6: "IT-1-6",
            7: "IT-2-1", 8: "IT-2-2", 9: "IT-2-3", 10: "IT-2-4", 11: "IT-2-5",
            16: "IT-3-4", 17: "IT-3-4", 18: "IT-3-6",
            19: "IT-4-1", 20: "IT-4-2", 21: "IT-4-3", 22: "IT-4-5", 25: "IT-4-6"}
DEFERRED = {12: "IT-2-6", 13: "IT-3-1", 14: "IT-3-2", 15: "IT-3-3", 23: "IT-4-4", 24: "IT-4-6"}
PI_AFTER_DEFERRAL = {12: None, 13: "PI-5", 14: None, 15: "PI-5", 23: None, 24: "PI-5"}
ENGAGED_AT_START = {"PI-1": range(1, 7), "PI-2": range(7, 13), "PI-3": range(13, 19), "PI-4": range(19, 25)}      # ft-25 comes later
TEAM = {1: "team-1", 2: "team-1", 3: "team-2", 4: "team-2", 5: "team-3", 6: "team-3",
        7: "team-1", 8: "team-2", 9: "team-3", 10: "team-1", 11: "team-2", 12: "team-3",
        13: "team-1", 14: "team-2", 15: "team-3", 16: "team-1", 17: "team-2", 18: "team-3",
        19: "team-1", 20: "team-2", 21: "team-3", 22: "team-1", 23: "team-2", 24: "team-3", 25: "team-1"}
PI_OF_FEATURE = {n: pi for pi, numbers in ENGAGED_AT_START.items() for n in numbers} | {25: "PI-4"}
LONG_STATUS = {"PI-1": "testing", "PI-2": "analyzing", "PI-3": "doing", "PI-4": "ready2deploy"}      # 2 days, every other status 1 day
CHAIN = ["analyzing", "ready4PI", "doing", "ready2test", "testing", "ready2deploy", "deploying"]

# --- IterationFeatures of the ART: Engaged, Added, Deferred, Released, ReleasedAdded, CarryOver -------------------------------------
ART_ROWS = {
    "IT-1-1": (6, 0, 0, 1, 0, 5), "IT-1-2": (5, 0, 0, 1, 0, 4), "IT-1-3": (4, 0, 0, 1, 0, 3),
    "IT-1-4": (3, 0, 0, 1, 0, 2), "IT-1-5": (2, 0, 0, 1, 0, 1), "IT-1-6": (1, 0, 0, 1, 0, 0),
    "IT-2-1": (6, 0, 0, 1, 0, 5), "IT-2-2": (5, 0, 0, 1, 0, 4), "IT-2-3": (4, 0, 0, 1, 0, 3),
    "IT-2-4": (3, 0, 0, 1, 0, 2), "IT-2-5": (2, 0, 0, 1, 0, 1), "IT-2-6": (1, 0, 1, 0, 0, 0),
    "IT-3-1": (6, 0, 1, 0, 0, 5), "IT-3-2": (5, 0, 1, 0, 0, 4), "IT-3-3": (4, 0, 1, 0, 0, 3),
    "IT-3-4": (3, 0, 0, 2, 0, 1), "IT-3-5": (1, 0, 0, 0, 0, 1), "IT-3-6": (1, 0, 0, 1, 0, 0),
    "IT-4-1": (6, 0, 0, 1, 0, 5), "IT-4-2": (5, 1, 0, 1, 0, 5), "IT-4-3": (5, 0, 0, 1, 0, 4),
    "IT-4-4": (4, 0, 1, 0, 0, 3), "IT-4-5": (3, 0, 0, 1, 0, 2), "IT-4-6": (2, 0, 1, 1, 1, 0)}
# team -> iteration -> (Engaged, Added, Deferred, Released) for PI-3 and PI-4 (t1 owns ft-13,16,19,22,25; t2 ft-14,17,20,23; t3 ft-15,18,21,24)
TEAM_ROWS = {
    "team-1": {"IT-3-1": (2, 0, 1, 0), "IT-3-2": (1, 0, 0, 0), "IT-3-3": (1, 0, 0, 0), "IT-3-4": (1, 0, 0, 1), "IT-3-5": (0, 0, 0, 0),
               "IT-3-6": (0, 0, 0, 0), "IT-4-1": (2, 0, 0, 1), "IT-4-2": (1, 1, 0, 0), "IT-4-3": (2, 0, 0, 0), "IT-4-4": (2, 0, 0, 0),
               "IT-4-5": (2, 0, 0, 1), "IT-4-6": (1, 0, 0, 1)},
    "team-2": {"IT-3-1": (2, 0, 0, 0), "IT-3-2": (2, 0, 1, 0), "IT-3-3": (1, 0, 0, 0), "IT-3-4": (1, 0, 0, 1), "IT-3-5": (0, 0, 0, 0),
               "IT-3-6": (0, 0, 0, 0), "IT-4-1": (2, 0, 0, 0), "IT-4-2": (2, 0, 0, 1), "IT-4-3": (1, 0, 0, 0), "IT-4-4": (1, 0, 1, 0),
               "IT-4-5": (0, 0, 0, 0), "IT-4-6": (0, 0, 0, 0)},
    "team-3": {"IT-3-1": (2, 0, 0, 0), "IT-3-2": (2, 0, 0, 0), "IT-3-3": (2, 0, 1, 0), "IT-3-4": (1, 0, 0, 0), "IT-3-5": (1, 0, 0, 0),
               "IT-3-6": (1, 0, 0, 1), "IT-4-1": (2, 0, 0, 0), "IT-4-2": (2, 0, 0, 0), "IT-4-3": (2, 0, 0, 1), "IT-4-4": (1, 0, 0, 0),
               "IT-4-5": (1, 0, 0, 0), "IT-4-6": (1, 0, 1, 0)}}

# --- Feature Predictability per PI -------------------------------------------------------------------------------------------
PLANNED = {"PI-1": 6, "PI-2": 6, "PI-3": 6, "PI-4": 6}
DONE_A = {"PI-1": 6, "PI-2": 5, "PI-3": 3, "PI-4": 4}          # released among the Features engaged at the start of the PI
DONE_B = {"PI-1": 6, "PI-2": 5, "PI-3": 3, "PI-4": 5}          # + ft-25, added during PI-4
CURRENT_PI = {"PI-1": 6, "PI-2": 5, "PI-3": 3, "PI-4": 5}      # what reading the PI field as it is today would count as planned

# --- the flow metrics of the ART (a Feature counts for the ART) --------------------------------------------------------------
VELOCITY = {label: sum(1 for released in RELEASED.values() if released == label) for label in ITERATION_STARTS}
# the 3 sub-tasks of the Program (TTT-26, 27 done in IT-2-1; TTT-28 in IT-3-4) are not Features: team "ART (other types)", counted by no ART figure
OTHER_TYPES_DONE = {"IT-2-1": 2, "IT-3-4": 1}
SUBTASKS = ("TTT-26", "TTT-27", "TTT-28")
SUBTASK_ITERATION = {"TTT-26": "IT-2-1", "TTT-27": "IT-2-1", "TTT-28": "IT-3-4"}
# the population by account (every ticket in exactly one): the 19 released Features, the 3 sub-tasks, the 6 deferred Features (open: a Backlog status)
ACCOUNTS = {"Done in an iteration": 19, "Done outside the iterations": 0, "Left out: not a Feature": 3, "Removed": 0, "Open": 6}
FLOW_TIME_DAYS = 8                                              # 7 statuses of 1 day, one of them lasting 2 days
FLOW_EFFICIENCY = {"PI-1": 5 / 8, "PI-2": 5 / 8, "PI-3": 5 / 8, "PI-4": 4 / 8}      # Active: analyzing, doing, testing, deploying

# --- where the cycle time goes: the Features RELEASED in each PI (by release date), time in each Active / Wait status -----------------
# ft-1..6 in PI-1, ft-7..11 in PI-2 (ft-12 is deferred), ft-16, 17, 18 in PI-3, ft-19..22 and ft-25 in PI-4 (ft-25 carries PI-4 and is released in IT-4-6).
PI_ITEMS = {"PI-1": 6, "PI-2": 5, "PI-3": 3, "PI-4": 5}
# average days per item in each status: 1 day, 2 days for the long status of the PI (the same for every item of the PI): 8 days in all
STATUS_AVERAGE_DAYS = {pi: {status: 2.0 if status == LONG_STATUS[pi] else 1.0 for status in CHAIN} for pi in PIS}
# share of the cycle time: the long status 2/8 = 25 %, each of the six others 1/8 = 12.5 %
STATUS_SHARE = {pi: {status: 0.25 if status == LONG_STATUS[pi] else 0.125 for status in CHAIN} for pi in PIS}


# DC28 (AL, 2026-11-18): the Features that took each move of the workflow. Every Feature (the 25: released ones and deferred ones, which go on to doing) was
# created in todo and moved analyzing -> ready4PI; the sub-tasks go todo -> released.
PLANS_TAKING_THE_MOVES = {"analyzing>ready4PI": 25}
# The 3 sub-tasks go todo -> analyzing -> doing -> released: analyzing -> doing and doing -> released are no move of the workflow (2 moves each, DC28).
DC28_SUBTASKS = ("TTT-26", "TTT-27", "TTT-28")
DC28_BAD_MOVES_PER_SUBTASK = 2


# --- the same scenario with no start status and no Wait status (AL, 2026-11-22: "flow metrics are still not computable") ----------------------
# Jira refuses the workflows (no project administrator right: no start status from Jira) and the sheet Workflows declares nothing (no IsStart, no Wait).
# Velocity does not change; Flow Time, Flow Efficiency and Flow Load are not computable in every iteration (never a 0); each of the 19 released Features
# is left out of Flow Time with the reason below; and the workbook says why, with its codes. Typed by hand.
NO_START_RELEASED = 19
NO_START_REASON = "No start status in the rules"
NO_START_HEALTH = ("No start status: Feature (no Flow Time, Flow Efficiency or Flow Load): set IsStart = 1 on the first in-progress status in "
                   "Workflows_from_user",
                   "No Wait status (no Flow Efficiency): set Category = Wait on the statuses where the work waits in Workflows_from_user")
NO_START_WARNINGS = ("J2D-JIRA-030 | the workflows of Jira could not be read",
                     "J2D-FLOW-010 | no start status in the Feature workflow: Flow Time, Flow Efficiency and Flow Load cannot be computed",
                     "J2D-FLOW-011 | no status is declared Wait: Flow Efficiency cannot be computed")
NO_START_CONTROLS = {"Every workflow kind with tickets has a start status (IsStart)": (0, 1, "REVIEW"),
                     "A status is declared Wait (Flow Efficiency)": (1, 0, "REVIEW")}

# --- FlowTimeCheck (AL, 2026-11-22): why the Flow Time of each iteration of the ART is what it is. Every released Feature is counted (it went through
# analyzing, a start status), nothing is left out; an iteration with no release says it. Without a start status (the variant above) every released
# Feature is left out for that reason. Typed by hand from VELOCITY.
FLOW_TIME_CHECK = {label: (v, 0, 0, 0, 0, v, 0, "" if v else "no item done in the iteration") for label, v in VELOCITY.items()}
FLOW_TIME_CHECK_NO_START = {label: (v, 0, v, 0, 0, 0, 0, f"all {v} done items left out: {v} no start status in the rules (set IsStart in Workflows)"
                                    if v else "no item done in the iteration") for label, v in VELOCITY.items()}

# --- FlowEfficiencyCheck (AL, 2026-11-22): each released Feature spends FLOW_TIME_DAYS days between its start and its release, of which the share
# FLOW_EFFICIENCY of its PI active, the rest waiting; the 3 Wait statuses of the sheet Workflows are exceptions of the Feature workflow
# (3 rows of StatusRules). An iteration with no release has nothing counted. Without a Wait status (the variant), every row says to declare one.
WAIT_RULES = 3
NO_WAIT_REASON = "no status is declared Wait (set Category = Wait in Workflows)"
NOTHING_COUNTED_REASON = "no item counted in Flow Time (see FlowTimeCheck)"
FLOW_EFFICIENCY_CHECK = {label: (v, 0, v * FLOW_TIME_DAYS * FLOW_EFFICIENCY["PI-" + label[3]], v * FLOW_TIME_DAYS * (1 - FLOW_EFFICIENCY["PI-" + label[3]]),
                                 WAIT_RULES, FLOW_EFFICIENCY["PI-" + label[3]], 0, "") if v else (0, 0, 0, 0, WAIT_RULES, "", 0, NOTHING_COUNTED_REASON)
                         for label, v in VELOCITY.items()}
FLOW_EFFICIENCY_CHECK_NO_START = {label: (0, 0, 0, 0, 0, "", 0, NO_WAIT_REASON) for label in VELOCITY}

# --- FlowLoadCheck (AL, 2026-11-22): the 24 iterations of 21 days are all measured (AS_OF is the end of the last one), the Feature workflow has one
# start status (analyzing, the first in progress of the workflow of Jira). Without a start status every row says to declare one. The loads
# themselves are compared with the reference by flow_excel.feature; here: what tells why.
NO_START_STATUS_REASON = "the workflow of the team has no start status (set IsStart in Workflows)"
FLOW_LOAD_CHECK = {label: (21, 21, 1, 0, "") for label in VELOCITY}                          # Days, DaysMeasured, StartStatuses, Gap, Reason
FLOW_LOAD_CHECK_NO_START = {label: (21, 21, 0, 0, NO_START_STATUS_REASON) for label in VELOCITY}

# --- FlowDistCheck (AL, 2026-11-22): every Feature of the scenario has the work type Feature: the released ones are all Feature, nothing is
# Unmapped. The same with or without a start status (the distribution does not need one).
FLOW_DISTRIBUTION_CHECK = {label: (v, (v, 0, 0, 0, 0), 0, "" if v else "no item done in the iteration", "") for label, v in VELOCITY.items()}
