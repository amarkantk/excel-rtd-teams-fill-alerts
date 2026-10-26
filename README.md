# Excel RTD → Teams Fill Alerts

Sends a Microsoft Teams message whenever a futures position changes, straight from Excel.
Everything is in one macro-enabled workbook, **TeamAlertFillsNotifier.xlsm**. You don't need an add-in, and nobody needs to edit code.

Each Teams card lists every contract that changed in the last few seconds, for example:

> **Position changed**
> TeamAlertFillsNotifier.xlsm · rtd_data[net_position] · 23-Oct-2026 20:46:46 · 2 change(s)
> ER3 SEP28 — +1 
> I SEP28 —  -1

---

## Contents

| File | What it is |
|---|---|
| `TeamAlertFillsNotifier.xlsm` | The workbook. All code is inside it. |
| `src/TeamsNotifier.bas` | Alert engine: watches the table, batches changes and posts to Teams |
| `src/clsAppEvents.cls` | Excel event hook used by TeamsNotifier |
| `src/Module1.bas` | Button macros (`TA_Start`, `TA_Stop`, `TA_Test`) that read the settings from `exch_map` |
| `src/frmSettings_code.txt` | Optional settings form (UserForm `frmSettings`) |
| `src/ThisWorkbook_code.txt` | Workbook open/close hooks |
| `README.md` | This guide |

The files in `src/` are copies of the code inside the workbook, kept as text so changes show up in version control.

---

## How the workbook works

| Sheet | Purpose |
|---|---|
| **Summary** | Count of products, watched products, open positions and net position for each group. It also has the **Start alerts / Stop alerts / Send test** buttons. |
| **data** | Table `rtd_data`, with one row per contract: Product, Group, Type, Exch, Watch, Inst, **net_position**, TT_Exch, TT_Symbol, TT_Expiry. C2 shows the feed status. |
| **exch_map** | **The only sheet you edit.** It holds the product setup table and the alert settings. |
| **contracts** | A single formula that builds the contract list from `exch_map`. You don't edit this sheet. |
| **How to use** | A short guide inside the workbook. |

### Product setup (table `exch_map`)

| Column | Meaning |
|---|---|
| Group | Product code, e.g. `ER3`, `SR3`, `FGBL`. Inter-product spreads use `A|B`, e.g. `I|ER3`. |
| Exch | NG exchange name, e.g. `ICE_L`, `CME`, `Eurex`, `ASX`, `CFE` |
| num_contract | How many contracts to list |
| Watch | `Y` = subscribe and alert, `N` = skip. A skipped group gets no RTD call and no alerts. |
| Type | `Serial` = the next N months. `Quarterly` = the next N Mar/Jun/Sep/Dec months. `Serial+Quarterly` = the next 4 serial months plus N quarterly months. |
| TT_Symbol | Product symbol on TT, used only when the server is TT. Blank means the group isn't looked up on TT. |

The contract list starts from the current month and rolls forward by itself.
To change the number of serial months, edit `contracts!B3` (default 4). To add a product group, add a row to the table.
`data!F2` warns you if the list outgrows the 400-row `rtd_data` table.

### Alert settings (`exch_map!I1:J11`)

| Cell | Setting | Default |
|---|---|---|
| J2 | **Teams webhook URL**: each person enters their own | *(blank)* |
| J3 | Recipient email. Use this only for a shared desk workflow (section 4). | *(blank)* |
| J4 | Table name | `rtd_data` |
| J5 | Column to watch | `net_position` |
| J6 | Row label column | `Product` |
| J7 | Batch seconds | `10` |
| J8 | Minimum change (0 = every fill) | `0` |
| J9 | Card title | `Position changed` |
| J11 | **RTD server**: `NG` or `TT` | `NG` |

Clicking **Start alerts** or **Send test**, or opening the workbook, copies these cells into the alert engine, so the sheet always takes priority.

### NG or TT RTD server

| | NG (`ngmd.rtdserver`) | TT (`tt.rtd`) |
|---|---|---|
| Instrument | `RTD(…,"INSTRUMENT",Exch,Product)` | `RTD("tt.rtd",,"Instr",TT_Exch,TT_Symbol,"FUTURE","Jun27")` |
| Position | `RTD(…,"POSITION",Inst,"net_position")` | `RTD("tt.rtd",,Inst,"NetPos")` |
| Feed status (data!C2) | Live status | "TT server (no status topic)" |

On TT, the exchange name is the same as Exch, except `ICE_L` becomes `ICE`. Check the TT_Symbol values against TT's own names before you rely on them.
TT inter-product spreads are not looked up yet.

---

## 1. First-time setup (each person)

1. Save the workbook to your own folder. Right-click it, choose **Properties**, tick **Unblock**, then click **OK**.
2. Open it and click **Enable Content**. If your firm only allows signed macros, ask IT to sign or approve the file.
3. Make sure your RTD server is running and logged in, either **ngmd.rtdserver** or the **TT Excel add-in**.
   - On NG, `data!C2` should show that the feed is pushing.
   - On TT, choose **TT** in `exch_map!J11`.
4. Create your Teams workflow (section 3) and paste the link into **`exch_map!J2`**.
5. On **Summary**, click **Send test**. When the card arrives in Teams, click **Start alerts**.

> **Avoid double alerts:** don't load the old `TeamsAlerts.xlam` add-in at the same time. Also remove TeamsNotifier code from any other workbook you have open.

### If the RTD cells show text starting with `#=`
Formulas sometimes get written into the workbook as text. To turn them back into live formulas:
1. Press **Ctrl+H**.
2. Set **Find** to `#=` and **Replace with** to `=`.
3. Set **Within** to **Workbook**, then click **Replace All**.

---

## 2. Maintainers: code layout inside the workbook

Press **Alt+F11**. The VBA project should contain:

| Folder | Item | Source |
|---|---|---|
| Microsoft Excel Objects | ThisWorkbook | `src/ThisWorkbook_code.txt` (its `Workbook_Open` calls `TA_PushSettings`) |
| Forms | frmSettings | `src/frmSettings_code.txt`. The UserForm's **(Name)** must be `frmSettings`. |
| Modules | TeamsNotifier | `src/TeamsNotifier.bas` |
| Modules | Module1 | `src/Module1.bas` (`TA_PushSettings`, `TA_Start`, `TA_Stop`, `TA_Test`) |
| Class Modules | clsAppEvents | `src/clsAppEvents.cls` |

Each Summary button runs one macro. To check, right-click the button and choose **Assign Macro**:

| Button | Macro |
|---|---|
| Start alerts | `TA_Start` |
| Stop alerts | `TA_Stop` |
| Send test | `TA_Test` |

After you change any code, run **Debug → Compile VBAProject**. Then export the changed modules back into `src/`.

---

## 3. Create your Teams workflow (about 5 minutes)

### Option A – alerts to a team channel (easiest)
1. In Teams, hover over the channel, click **•••**, then choose **Workflows**.
2. Choose **"Send webhook alerts to a channel"**, click **Next**, then click **Add workflow**.
3. Click **Copy webhook link** and paste it into `exch_map!J2`.

### Option B – alerts just to you (personal chat)
The template "Send webhook alerts to a chat" needs **three fixes**. Without them, it posts to your self-chat, which Teams rejects with *"Call made for a thread which is not a ChatThread"*.

1. In Teams, open **Apps → Workflows → Create**. Choose **"Send webhook alerts to a chat"** and pick any chat. Click **Add workflow**, then **Copy webhook link**.
2. Go to **Workflows → Home**, open the workflow and click **Edit**. Inside **Attachments is null**, expand **2 Cases**.
3. Under the **False** branch, open **For each → Post card in a chat or channel** and set these fields:
   - **Post as:** Flow bot
   - **Post in:** Chat with Flow bot
   - **Recipient:** your work email
   - **Adaptive Card:** click **fx** and enter `items('For_each')?['content']`
4. Under the **True** branch, open **Post card in a chat or channel 1**. Set the same fields, but for **Adaptive Card** enter `string(variables('Body'))`.
5. Click **Save**. Alerts arrive in a Teams chat called **Workflows**; pin it so it's easy to find.

> If `items('For_each')` is rejected, open the **For each** step in Code view and use the name shown there, with underscores in place of spaces.

---

## 4. (Optional) One shared workflow for the whole desk

1. Build Option B. In **both** "Post card" steps, set **Recipient** to `triggerBody()?['recipient']`.
2. Add one or two co-owners on the workflow page.
3. Each colleague puts the shared URL in `exch_map!J2` and **their own email** in `exch_map!J3`.

If J3 is left as someone else's email, that person receives the alerts. With your own workflow, J3 is ignored.
The workflow owners can also see everyone's messages in its run history, so check with compliance before you choose this.

---

## 5. Sharing the workbook

Before you send the file to anyone or commit it:

- **Clear `exch_map!J2` (webhook URL) and `J3` (email).** Anyone who has the URL can post to your chat.
- **Leave no live positions in the file.** Save it with the feed off, or clear the values first.
- **Send only the `.xlsm`.** Colleagues follow section 1; they don't need to import any code.
- Settings are stored per Windows user. Each person's clicks copy their own `exch_map` values into the alert engine.

---

## 6. Troubleshooting

| Symptom | Fix |
|---|---|
| *"Enter your Teams webhook URL first (Teams Alerts > Settings)"* | The button is running the old add-in macro. Right-click the button, choose **Assign Macro**, then pick `TA_Start` or `TA_Test`. Also untick the TeamsAlerts.xlam add-in if it's loaded. |
| *"Paste your Teams webhook URL in exch_map!J2 first"* | J2 is empty or doesn't start with `https://`. |
| Inst / net_position show `#=…` text | Do the Ctrl+H step in section 1. |
| Inst shows "Instrument Not Found" | Check the Group and Exch (NG), or TT_Symbol (TT), in `exch_map`. |
| `data!F2` says "TABLE TOO SMALL" | Extend the `rtd_data` table; its formulas copy down automatically. |
| "SEND FAILED" | Check the URL. The engine retries every 60 s and keeps the changes until they're sent. |
| HTTP 401 / 403 | In the workflow trigger, set **Who can trigger the flow?** to **Anyone**. |
| Status says "sent" but nothing arrives in Teams | Open **Workflows → your workflow → Run history** and read the step marked in red. |
| Every fill alerts twice | Another copy of TeamsNotifier is running, either the add-in or another workbook. Remove it. |
| Alerts stopped after a VBA error | Click **Start alerts** again. |
| Compile error on `StopMonitoring` | Change `TA_Stop` in Module1 to the Stop sub name used in TeamsNotifier. |

### How it behaves
- **Only numeric changes alert.** Blanks, `#N/A` and "loading" text are ignored. net_position shows blank (not 0) when the feed has no number, so a glitch never looks like a position going flat.
- **The first 30 seconds after Start are quiet.** Positions are learned without sending alerts.
- **New and removed contracts.** A contract with a non-zero position that appears is reported as `(new) → value`. One that disappears, for example after the monthly roll, is reported as `value → (removed)`.
- **Switching servers.** Click **Stop alerts** before you switch between NG and TT, then click **Start alerts** again.
- **Excel must stay open.** Excel needs to be running with the workbook open and the feed connected. Nothing is sent while Excel is closed.

> Position and fill data is sensitive. Send it only to approved chats or channels, and keep the webhook URL private.
