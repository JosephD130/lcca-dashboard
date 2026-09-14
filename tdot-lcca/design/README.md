# Design canvas: TDOT LCCA Workbook UI

Source for the published design canvas at
https://claude.ai/code/artifact/c57d7afb-262b-4884-986f-f66f9b84b0cf

Six artboards showing the workbook's screens as mockups, so the look and the first-run
experience can be reviewed without opening Excel:

| File | Artboard |
|---|---|
| `Main.dc.html` | General Information, v1.2.0 (how-to card, live input checklist, section bands, buttons). Carries a switch between an empty project and the Outlaw Field test project. |
| `Summary.dc.html` | Summary sheet, v1.2.0 (results table, verdict line, RealCost-style comparison block, five charts) |
| `AltSheet.dc.html` | An alternative worksheet, v1.2.0 (navigation bar, input guide, the Price from control, pay-item table) |
| `TypicalValues.dc.html` | The new Typical Values reference sheet (unit-cost section) |
| `BeforeGeneralInfo.dc.html` | General Information as it is in v1.1.2 |
| `BeforeSummary.dc.html` | Summary as it is in v1.1.2 (five columns and one chart) |
| `canvas.json` | Layout, artboard titles and the sticky notes |

Figures are from the Outlaw Field (CKV) Runway 17-35 test run, so they match
`verification/new_project_CKV`. Colours and type are lifted from the workbook itself:
Arial, `#D9D9D9` input fill, `#1D2733` buttons, `#2A78D6` accent, `#EAF2FB` bands,
and the tab colors now set on the sheets. The TDOT logo is a placeholder block; the
workbook carries the real image.

`tdot-lcca-workbook-ui.html` is the published page (generated; regenerate by re-seeding
from the files above rather than editing it).
