# Tabroom import checks

Eden Prairie Eagle Invitational, tournament 41328, is covered by the automated tests:

- `Fixtures/tournament-41328.json`
- `Fixtures/events-41328.json`
- `Fixtures/senate-464773.json`
- `Fixtures/vld-464778.json`

`TabroomParserTests` checks that Senate finalists score separately from prelims, and that Varsity LD with no elimination break uses placement points.
