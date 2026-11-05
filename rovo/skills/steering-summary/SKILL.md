---
name: steering-summary
description: Write a short steering note for the RTE from the Confluence page "Flow metrics <PI>". It gives signals, risks and suggested actions per team, each tied to a figure on the page or a JQL link. Use before an iteration review, an Inspect and Adapt, or a PI planning. A proposal, never a decision.
---

# Steering note for the RTE

1. Read the page "Flow metrics <PI>": the header and tables A, B and C. Quote every value and trend the way the
   explain method does: exactly as written, with its period and team.
2. Pick 3 to 5 signals, the most important first, only from the page:
   - Flow Time P85 is higher than in the previous period;
   - Flow Load at the period end is above the team's Flow Load average for the period;
   - Flow Efficiency is down, or not computable;
   - in the distribution, the Defect share is up, or Unmapped is above 0 (the work type is missing);
   - a check of table C is KO with a Blocker severity.
3. For each signal, give:
   - the team;
   - the two values quoted with their periods;
   - one likely cause, written as a question;
   - one suggested action.
4. Then a section "Data to fix first": the KO checks that make a figure not computable or leave items out, with their
   Items.
5. Keep the note under 20 lines. Title it "Proposal for the RTE - <PI> - as of <date>".
6. Never give a decision. Never rank people or teams. Never use a figure that is not on the page.
