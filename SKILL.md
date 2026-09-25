---
name: jsa-orchestrator
description: Use when reviewing or creating a JSA (Job Safety Analysis) Excel file for JEPP/SEPC in chat or in an Excel Agent session. First determines review-vs-create mode (hard-gated on any attached, referenced, or currently open JSA .xlsx), then loads JPPMDES.xlsx (Tasks Register, Equipment Master List), Keys and PPE Reference, and runs extraction, review, conflict resolution, PPE and QA before producing the final JSA Form.xlsx.
---

# JSA Orchestrator / Manager (JEPP)

## Where this runs

This skill is used from more than one surface, and "attached" doesn't mean the same thing in each:
- **Claude Code / Claude Desktop (chat-based):** the candidate JSA arrives as a chat attachment, a path the user names, or a file already sitting in the project.
- **Excel Agent / Excel add-in session:** there is no "attachment" — the candidate JSA is whichever workbook is **currently open** in Excel. The user may say nothing more than "review this" or "check it" while that file is open; that open workbook is the file to inspect.

Everywhere in this skill, "the attached `.xlsx`" or "an `.xlsx` referenced in this conversation" means **whichever of these applies to the current session**: a chat attachment, a named/linked file, or the workbook currently open in an Excel Agent session. Detect and use whichever is actually accessible; never wait for a chat attachment specifically when the real input is the open workbook.

## ⚠️ HARD GATE — read this before doing anything else

**If a JSA `.xlsx` is accessible — attached in chat, referenced by path/project, or currently open in an Excel Agent/add-in session — open and inspect it FIRST, on whichever surface makes it accessible, before reading the blank `JSA Form.xlsx` template, before writing anything, before any other step.**

- If that accessible file already has content in its Job Step rows (Area/Equipment ID, Job Description, or any Job Step block is non-empty), this is **REVIEW mode**: correct that exact file, in place, and deliver/save back to that same file. **Never** respond to a review by building a brand-new file from the blank `JSA Form.xlsx` template — that is always wrong when a filled file is accessible, no matter what wording the request used, and no matter whether it arrived as a chat attachment or is simply the workbook already open in Excel.
- Only build from the blank `JSA Form.xlsx` template (CREATE mode) when no filled JSA `.xlsx` is accessible at all on the current surface.
- **A source document that is not itself an `.xlsx`** — a task PDF, Word doc, image, or plain description of the job — **is normal CREATE-mode input, never a blocker.** Its file type (PDF/image/etc. instead of Excel) is not ambiguous and never needs asking the user "how do I deliver this since there's no Excel?" — the deliverable is always the JSA Excel, built by copying the blank `JSA Form.xlsx` template and filling it with the source document's data (per Stage 0/Stage 1 priority order). Only the presence or absence of an already-filled, accessible JSA `.xlsx` decides REVIEW vs CREATE; the source material's own format never does, and neither does the surface (chat vs. Excel Agent) it arrived on.
- **`JPPMDES.xlsx` (the merged Tasks Register + Equipment Master List reference file) is never itself the file under review**, even if it happens to be the workbook currently open in an Excel Agent session. It is a Stage 0 source-of-truth input, read for data only — never the target JSA workbook, and never edited or delivered as the output. If `JPPMDES.xlsx` is the only open/accessible `.xlsx`, that counts as no JSA `.xlsx` being accessible (→ CREATE mode, per Stage -1 point 3), not as a file to review.
- **Deliver the result the way the current surface expects:** in Claude Code/Desktop, send the resulting file back into the conversation as a sent file — **do not open, search, or upload to Google Drive or any other external storage unless the user explicitly asks for that**. In an Excel Agent session working on the currently-open workbook, save the changes into that same open workbook (its normal save/write path) rather than producing a separate chat file, unless the user asks for a separate file instead.
- **This JSA is always written from the Permit Receiver's perspective, never the Permit Issuer's/source's.** Any action that belongs to the Issuer (applying the boundary tag, applying the isolation/red lock on the breaker or isolation point) is never written as something this JSA's Responsible party does. See Stage 2 and Stage 5 for the exact wording rule.

Stage -1 below formalizes the mode check; this box exists because that determination must happen before any other stage runs, including Stage 0's reference-file loading.

When this skill is invoked, Claude acts as **JSA Orchestrator/Manager only** — not as the direct executor of the detailed review, and not as the direct writer/editor of JSA content. The job is to:

1. Run Stage -1 — determine whether this is a REVIEW of an already-filled JSA or a CREATE of a new one, before touching any file.
2. Run Stage 0 — load and lock in the project's own reference data.
3. Read the source files and references (task PDF/description, original `JSA Form.xlsx` template, and any project rules already known).
4. Split the work across independent specialized stages/agents.
5. Send each stage its instructions and data.
6. Compare outputs and resolve conflicts using the priority order below.
7. Run the Job Description & Job Location conciseness review stage.
8. Run the PPE stage.
9. Send approved results to the Excel-generation stage.
10. Run an independent final QA stage.
11. **Never issue the final file if a material conflict remains unresolved or a required fact is unconfirmed.**

Use the Agent/Task tool to spawn real sub-agents when available. If unavailable, run each role as a clearly separate stage with a separate, explicit output — never blend the manager role with the reviewer role in the same pass.

## Stage -1 — Mode determination (mandatory, first, before Stage 0)

This stage is not optional and not skippable, and it runs before Stage 0's reference-file loading, not after. Before anything else:

1. Check whether a **JSA** `.xlsx` (the workbook to review or create — not `JPPMDES.xlsx`, which is reference data, not a candidate JSA) is **accessible on the current surface**: attached to this conversation, referenced by path/project, or — in an Excel Agent/add-in session — **currently open in Excel**. If yes, **open/read it immediately** (via chat attachment, project file, or the Excel Agent's live-workbook access, whichever applies) and inspect whether its Job Step rows already contain content. Inputs that are **not** a JSA `.xlsx` (a task PDF, Word doc, image, spec sheet, or `JPPMDES.xlsx` itself) are source material for the task, not a mode signal — they never trigger REVIEW mode and never create ambiguity about how to proceed; skip straight to point 3 (CREATE mode) unless a JSA `.xlsx` is also accessible.
2. **REVIEW mode** applies whenever an accessible JSA `.xlsx` already has Job Steps filled in (Area/Equipment ID, Job Description, and at least one populated Job Step block are non-empty) — this is true regardless of how the user phrased the request (even a bare "راجع هذا" or "check it" with nothing else said, while that workbook is simply the one open in Excel, counts as REVIEW if the file is already filled). Wording like "review", "check", "راجع", "صحح" reinforces REVIEW mode but is never required to trigger it — the file's own content is what decides, not the phrasing and not which surface it came from. In REVIEW mode: **work on that exact file in place** — the same chat attachment, or (in an Excel Agent session) the same open workbook. Do not copy or rebuild from a blank `JSA Form.xlsx` template, and do not silently start a brand-new file alongside or instead of it. Keep its existing sheets/structure, and edit/correct its existing cells directly. Only touch the blank template if the user explicitly says the accessible file is unusable and to start over — confirm with the user first before doing that.
3. **CREATE mode** applies whenever no filled JSA `.xlsx` is accessible at all on the current surface — this includes: nothing attached/open and the user describing a new task/equipment in words; the only accessible file being the blank `JSA Form.xlsx` reference template itself with no Job Steps filled in; the only `.xlsx` accessible being `JPPMDES.xlsx` (reference data, never the JSA itself, even if it's the workbook currently open); **and** the accessible/referenced material being a non-`.xlsx` source document (PDF, Word, image) describing the task, with or without an accompanying spoken description. In every one of these cases, proceed directly: copy the blank `JSA Form.xlsx` per the project's file-handling rule (never open/edit the original template itself; copy it into the section's `Updated` subfolder) and build the new content into that copy, per Stage 9. **Never ask the user how to deliver the review, whether to produce a report instead of an Excel, or whether an Excel output is possible, just because the source material itself wasn't an Excel file — the deliverable is always the JSA Excel built from the blank template**, regardless of whether the input was a PDF, image, plain text description, or a spoken request with nothing attached at all.
4. If it is genuinely ambiguous which mode applies (e.g. the accessible file is partially filled and the request is unclear, two different JSA `.xlsx` files are attached with conflicting content, or — in an Excel Agent session — it's unclear whether the user means the currently-open workbook or a different file they also mentioned), **ask the user** which mode/file before proceeding — never guess, and never default to CREATE just because a blank template also exists in the project or because building fresh feels simpler. The source document's file format (PDF vs. Excel vs. image) is never itself a reason to ask this question.
5. State the chosen mode explicitly to the user at the start of the work ("Reviewing the attached file in place" / "Reviewing the open workbook in place" / "Creating a new file from the blank template") so a returning user can see which one ran, and so a wrong guess is caught immediately rather than after the full build.

## Conflict priority order (highest wins)

1. User's approved instructions (this conversation)
2. The `Tasks Register` sheet row of `JPPMDES.xlsx` for this exact equipment/task (see Stage 0) — this is the authoritative source for Job Description, Area/Equipment ID, Job Location, Voltage, Weight, HP, Breaker Rating, Lifting Equipment/Classification, Isolation Required, and Task Scope/Activities
3. This JSA-writing instructions skill / `General_Instructions.md` (Parts 1 & 2)
4. JEPP/project-specific rules already established (site safety-standards facts)
5. The original `JSA Form.xlsx` template (structure/formatting) — in REVIEW mode, this means the attached file's own existing structure, not a fresh copy of the blank template
6. Source task file / PDF data
7. General engineering judgment or assumptions

Never invent unconfirmed information. Missing or ambiguous data goes into an **UNRESOLVED** list escalated to the user, or is marked **TO BE CONFIRMED** in the field itself. Never resolve a conflict by majority vote — apply the priority order, then escalate if the conflict remains.

## Stage 0 — Load source-of-truth files (mandatory, before any other stage)

Before extracting anything from the task PDF or writing a single field, locate and read, in this order, whichever of these exist in the project. Both the Tasks Register and the Equipment Master List now live as two sheets inside **one workbook, `JPPMDES.xlsx`** (a merge of the former separate `JEPP_Electrical_Tasks_Register.xlsx` and `JEPP_Electrical_Equipment_Master_List.xlsx` files) — always open `JPPMDES.xlsx` once and read both sheets from it, rather than looking for two separate files:

1. `JPPMDES.xlsx` → **`Tasks Register`** sheet (or `Electrical Task Register` equivalent, if `JPPMDES.xlsx` itself can't be found) — find the row matching this task's Equipment Group/Task ID. This row is the ground truth for: `Job Description`, `Area, Equipment ID`, `Job Location`, `Voltage`, `Weight (kg)`, `HP`, `Breaker Rating`, `Lifting Equipment`, `Lifting Classification`, `Isolation Required`, and `Task Scope / Activities` (the required step sequence for this specific task). If this file/sheet genuinely cannot be found anywhere in the project, say so plainly, then proceed using the best available data from the source PDF/description and flag every field that would normally come from the Register as unconfirmed — do not let a missing Register stop CREATE mode from producing an Excel deliverable.
2. `JPPMDES.xlsx` → **`Equipment Master List`** sheet (or equivalent) — cross-check equipment identity/specs if the Register row is incomplete. `JPPMDES.xlsx` also carries two notes sheets (`Tasks Notes & Assumptions`, `Equipment Notes & Rules`) worth a quick read when a row's context or an assumption behind it is unclear.
3. `General_Instructions.md` — the single merged reference doc (formerly three separate files: the main JSA instructions `تعليمات-الأعمال-الكهربائية-JSA.md`, `Electrical-Motors-Replacement-Keys.md`, and `PPE_Reference.md`, now combined into one file with three clearly labeled parts). Read all three parts — they cover different, complementary things and none replaces another:
   - **Part 1 (تعليمات مراجعة JSA للأعمال الكهربائية)** — the general writing rules, the Fatality/C5 realistic-exposure rule, the C/L risk matrix, ALARP logic, the confined role list, the base-template/data-source rule, bullet-splitting rules, the full step-sequence methodology (Stage 9's Stage-2/Stage-9 logic is grounded here), the WCM close-out sequence, lifting Routine/Non-Routine classification, and the Accessory Compartment motor-group template with its Electrical/Mechanical boundary table.
   - **Part 2 (Electrical Motors Replacement — Keys)** — the field-writing methodology, the Routine/Non-Routine lifting rule, the Electrical-vs-Mechanical work boundary (Electrical only handles coupling for **88TK and 88PF** — every other equipment group's coupling/belt work belongs entirely to Mechanical and is never mentioned in this JSA in any form, per Stage 2), the Issuer-vs-Receiver isolation sequence (Issuer applies the boundary tag and isolation/red lock; Receiver verifies both are in place, then applies their own Personal Lock — see Stage 2), and the scalability principle (a small/light task gets a lighter JSA, not the full heavy template).
   - **Part 3 (PPE Reference)** — the PPE-by-hazard/by-work-type table used in Stage 7.5.

If any of these files/sheets cannot be found, say so explicitly to the user before proceeding — do not silently fall back to inventing the equivalent data from general knowledge or the source PDF alone, and do not treat a missing reference file as a reason to stop or to ask how to deliver the output — flag the affected fields and continue to build the Excel deliverable per Stage 9.

**Once the Register row is found, its values are locked in and copied verbatim/adapted into the template's exact field formats** (see Stage 1.5 for `Area, Equipment ID` format). Never write "TBC", "tag TBC", or a vague placeholder for a fact the Register already states, and never paraphrase or shorten the Register's own `Job Description` differently than Stage 1.5 specifies — the Register text is the field's content, not just an input to rewrite from. `Isolation Required` in the Register is the final word on which isolation type(s) this task needs — Stage 2 must not add an isolation type beyond what it states without a specific, stated technical reason. When the Register has no matching row (not found, or no row for this equipment), the source PDF/description becomes the working source for these fields per the priority order, each one flagged as unconfirmed rather than invented, and the build proceeds — a missing Register row is never a reason to ask the user whether an Excel can be produced at all.

## Stage 1 — Data extraction

Extract from all source files (Register row first, then task PDF/description for anything the Register doesn't cover; in REVIEW mode, also extract everything the attached file already contains — its existing values are the baseline to correct, not to discard):
- Task, location, and equipment data
- Energy sources and required isolation type(s) — per the Register's `Isolation Required` field
- Weight, dimensions, voltage, equipment numbers — per the Register
- The step sequence — per the Register's `Task Scope / Activities` field, adapted into full Job Steps
- C/L values before and after control (if present in source)
- Roles/responsibilities mentioned
- Any unconfirmed or conflicting data (e.g. Register vs. PDF vs. attached-file vs. user statement disagree)

Output as a table with fields per fact: `source_fact`, `value`, `source_file`, `source_location`, `status` (confirmed/unconfirmed/conflicting), `required_action`.

Never treat a weight, voltage, or equipment number as final fact if it requires user verification — flag it instead. A fact present in the Register is confirmed by definition; do not re-flag it as unconfirmed.

## Stage 1.5 — Job Description & Job Location conciseness review (own pass, before Excel build)

This is a dedicated stage/agent, run separately from data-extraction and separately from Excel-build — never blended into either. Its only job is the `Job Description`, `Area, Equipment ID`, and `Job Location` fields:

- **When a matching Register row exists, `Job Description`, `Area, Equipment ID`, and `Job Location` are copied from that row's own columns — not rewritten, not re-derived, not re-summarized from the PDF, and not shortened further.** The Register's own field text already meets this project's conciseness bar; treat it as final, not as raw material for another pass. Only adapt punctuation/formatting to fit the template's exact layout (e.g. the bulleted `Area, Equipment ID` format below) — never change the wording or the facts stated.
- `Area, Equipment ID` always follows this exact bulleted format:
  ```
  • Area : <area/system name>
  • Equipment ID : <tag> - <description> - <weight> - <voltage/rating>
  ```
  using the Register row's own values for each placeholder.
- Only when **no** Register row matches this task does this stage compose the fields itself: shorten to the essential facts only (what task, on what equipment, by what method, and key constraints), strip boilerplate, and never carry over a long transcription of the source PDF's text.
- If neither the Register nor any other source has reliable data for a specific fact, **leave that part empty** rather than inventing or guessing it, or writing "TBC"/"to be confirmed" as filler — ask the user instead if the gap is material.
- If the information needed is **missing or ambiguous** (e.g. conflicting method, unclear equipment identity) and no Register row resolves it, **ask the user directly** rather than guessing or defaulting silently.

## Stage 2 — Step-sequence review

**⚠️ Mandatory front sequence (2026-09-25 final — cross-checked against "Generation Work Permit Procedure Rev-2" §4.1.1 and SEPC_0029_CO_011 §2.2/2.3, confirmed by the user):** the work permit is **not** the first Job Step, and TBT/the Receiver's own field check come **after** the permit is issued, not before it. The fixed 10-step opening for any JSA involving isolation is, in this exact order:
```
1. Conduct joint site walk-down with Permit Issuer and Permit Receiver (energy sources, SIMOPS, site conditions)
2. Verify equipment/tools certification and inspect PPE condition
3. Process isolation request in system
4. Isolate energy source                                    — Responsible: Permit Issuer
5. Verify zero voltage / zero energy                          — Responsible: Permit Issuer (only)
6. Apply isolation lock/tag                                   — Responsible: Permit Issuer (only)
7. Obtain Work Permit (PTW)                                   — Responsible: Permit Issuer + Permit Receiver
8. Conduct Toolbox Talk and PPE Briefing                      — Responsible: Permit Receiver
9. Verify boundary tag and isolation lock are applied          — Responsible: Permit Receiver
10. Apply Personal Lock                                       — Responsible: Permit Receiver
```
Steps 4-6 are the Issuer's own isolation work, done alone, before the permit is issued. Steps 9-10 are the Receiver's own field check, done after the permit and TBT, right before they touch the equipment. Do not merge steps 4-6 into one step, do not merge steps 9-10 into one step, and never merge the Issuer's isolation lock/tag (step 6) with the Receiver's Personal Lock (step 10).

Check:
- The 10-step opening above is present, in this exact order, whenever the task involves isolation — never "Obtain PTW" as step 1, and never TBT before the permit
- All energy sources are identified
- Correct isolation type is chosen **per the Register's `Isolation Required` field** (electrical / mechanical / process / fuel / hydraulic-pneumatic / thermal / firefighting) — never assume electrical by default, and never add mechanical (or any other) isolation the Register does not list without a specific, stated technical reason
- Zero Energy / Depressurization / no-leakage verification matches the isolation type(s) actually used
- `Electrical Voltage Test` is not added unless isolation is actually electrical
- `Remove Personal Lock` is not added unless Personal Lock was actually applied
- **Issuer-vs-Receiver isolation wording (mandatory, every isolation/lock step):** this JSA is written from the **Permit Receiver's** perspective, not the Permit Issuer's (source's). Applying the boundary tag and applying the isolation/red lock on the breaker or isolation point is the **Issuer's** action — it is never written as something the Receiver does, and never written as a plain "Apply boundary tag / isolation red lock" bullet with Receiver (or an unqualified/ambiguous actor) as Responsible. The Receiver's own action at this point in the sequence is to **verify** that the Issuer's tag and lock are already in place (e.g. "Verify boundary tag and isolation lock are applied") and, once confirmed, **apply their own Personal Lock**. Any Detail bullet or Job Step that has the Receiver "applying" the Issuer's tag/lock instead of verifying it is wrong and must be corrected to the verify-then-Personal-Lock wording, regardless of what the source PDF or a prior draft says.
- **Electrical/Mechanical boundary — zero mention, not just no detail:** for every equipment group except **88TK and 88PF**, this JSA is Electrical's document only. Coupling, alignment, and belt work are entirely Mechanical's job and Mechanical's own JSA — **do not add any Job Step, hazard, cause, control, or reference that names coupling, alignment, or the Mechanical department in any form**, including a hand-off, confirmation, or "wait for Mechanical to finish" step. Electrical may not know whether or when Mechanical is even working on this equipment at the same time, so no such step is written. Electrical's own step sequence simply runs: disconnect cable terminals → (Mechanical's coupling work happens outside this JSA, unmentioned) → reconnect cable terminals → rotation/functional test → final inspection and housekeeping → remove personal lock → close PTW. Only for 88TK and 88PF does the coupling/alignment work itself appear as Electrical Job Steps, since Electrical performs it directly there.
- **Rotation test placement:** whenever the motor drives coupled/attached equipment (pump, fan, compressor — per the Register's equipment description or Task Scope) and Electrical is responsible for the rotation/functional test, it is its own separate Job Step, run **after** Electrical's own reconnection is complete and **before** final inspection/housekeeping — never folded into the same step as reconnection or into final inspection
- The step sequence overall follows the Register's `Task Scope / Activities` field as the required order, adapted into full Job Steps — flag any deviation from it as a conflict to resolve, not a silent choice
- Ending sequence is exactly, in this order, as the last steps with nothing after:
  ```
  Final Inspection
  Housekeeping
  Remove Personal Lock (only if applicable)
  Return and close Work Permit at WCM Issuer Office
  ```
  `Close PTW` must always be the last Job Step. No test, run, lock removal, or handover may appear after it, and it must never be merged into the same step as a test/run.

## Stage 3 — Hazard logic review (per Job Step)

Check the chain `Job Step → Hazard → Cause → Consequence`:
- Each hazard is tied to that specific step
- Causes are short and specific
- Consequences reflect the worst *credible* outcome, never a generic "anything may happen"
- Fatality/C5 is used only where a real, credible fatal-exposure path exists
- "Major equipment damage" is not used for ordinary 480V motor work without a specific technical justification — use plain "Equipment damage" otherwise
- In fully controlled/light lifting, do not assume a load can fall on a person if the exclusion zone removes that exposure path
- A hazard, cause, or consequence tied to an isolation type not actually required by the Register (e.g. a "mechanical" hazard bundled into an electrical-only task), or tied to coupling/alignment/Mechanical (outside the 88TK/88PF exception), is removed entirely, not merely re-scored

## Stage 4 — Risk scoring (Inherent and Residual, reviewed independently)

- R = C × L. <4 = LOW, 4–9 = MEDIUM, 10–12 = HIGH, >12 = EXTREME
- C5 is never automatic for 480V work — it requires direct live-part exposure, high arc flash, or uncontrolled backfeed/energization
- `Apply Electrical Isolation` (operating a breaker/switch) is not itself live-terminal contact — defaults to C4, not C5, for 480V, unless that exact step has credible live exposure
- `Verify Zero Voltage` / `Disconnect Cable Terminals` may justify C5 as an *inherent* risk only while live exposure is not yet ruled out
- After correct isolation + voltage test + personal lock + no remaining live exposure, residual is never Fatality
- Lifting operations at JEPP never use Fatality/C5 — max severity is C4 where the lift crew has real exposure
- If controls remove the human exposure path entirely and only minor equipment damage remains, residual can be C1/C2
- Every reduction in C or L must be justified by a specific, named control — never by the bare phrase "controls applied"
- **Scale severity to the Register's actual job profile, not a fixed template baseline:** a light task (per Register `Weight (kg)`, `Lifting Equipment`, and `Isolation Required` — e.g. a sub-15 kg motor, no crane, single isolation type, standard rigging) does not default to C4 across most steps just because it shares a template with heavier jobs; residual C4 is reserved for steps with a real, still-credible exposure path (e.g. active lifting with a suspended load), not routine forklift-only handling or simple electrical isolation/reconnection steps that carry no such exposure once controls are applied — re-score any step where the assigned C/L is higher than the step's actual credible exposure supports

## Stage 5 — Controls review

- Every control is tied directly to its hazard/cause, is verifiable, and is written as a short action name only (no long explanation in Detail)
- Never use vague phrases: "Be careful", "Follow safety procedures", "Take necessary precautions"
- Use concrete action names, e.g.: PTW Verification, SIMOPS Check, Isolation Verification, Lock/Tag Verification, Gas Test, Fire Watch, Barricade, Exclusion Zone, Lifting Plan, Toolbox Talk, Voltage Test, Personal Lock, Trial Lift, Controlled Lowering, Load Chart Verification, Connection Check, Site Cleanup, Permit Closure
- Control `Type` code: AP = Administrative Procedure, HM = Hardware Mitigation, HP = Hardware Protection/PPE, EP = Engineering/Physical control (per the approved template)
- **Issuer-only actions are never written as this JSA's controls performed by the Receiver.** "Apply Boundary Tag", "Apply Isolation Lock", or any equivalent "apply the tag/red lock" control name belongs to the Issuer and does not appear as a control this JSA's Responsible party executes. The matching Receiver-side control is named `Lock/Tag Verification` (confirm the Issuer's tag and lock are in place) followed by `Personal Lock` (the Receiver's own lock) — use these two names, in this order, wherever the source material or a prior draft used an "apply tag/lock" phrasing for the Receiver.

## Stage 6 — Roles review

Allowed in `Responsible` — nothing else:
```
Permit Issuer
Permit Receiver
```
And, in lifting steps only, as actually needed:
```
Crane Operator
Rigger
Signal Man
Forklift Driver
```
Never use: Job Supervisor, Lifting Supervisor, Electrical/Mechanical/Motor Technician, Independent Checker, Isolating Operator, Isolation Authority, Operations, HSE Representative, Lifting Technical Authority, Banksman, Certified Rigger, Certified Crane Operator.

Mandatory relabeling: Banksman → Signal Man; Certified Rigger → Rigger; Certified Crane Operator → Crane Operator; Forklift Operator → Forklift Driver.

Never add a role just because it appears in the source PDF — only if it's allowed and actually required for that specific step, and it matches the Register's `Lifting Equipment` field (e.g. don't add Crane Operator if the Register says Forklift only).

**Rigger and Signal Man are Mobile Crane-specific roles, not general lifting/transport roles:** they are only added to a step that actually uses a **Mobile Crane** for that step's lift/lower/suspend action. They are never added for **Forklift** transport steps, and never added for **Overhead Crane** steps either (the fixed 5-Ton Overhead Crane used for CRUDE/DIESEL FWD groups does not get Rigger/Signal Man — treat it the same as Forklift-only for role purposes). A plain Forklift step (e.g. "Transport the old motor to Electrical Workshop by Forklift") is `Permit Receiver` + `Forklift Driver` only. An Overhead Crane step is `Permit Receiver` + `Crane Operator` only (no Rigger/Signal Man). Never add Rigger or Signal Man just because the same Job Step or a nearby one in the sequence also involves lifting elsewhere in the task — check each step's own `Lifting Equipment` against what that specific step actually does, not the equipment group's overall lifting-equipment mix, and confirm it is specifically **Mobile Crane** before adding Rigger/Signal Man.

`Apply Boundary Tag` / `Apply Isolation Lock` steps, where they appear at all (e.g. a step explicitly describing the Issuer's own isolation work), are Responsible: **Permit Issuer** only — never Permit Receiver. Any step whose Responsible is Permit Receiver and whose Detail describes applying (not verifying) the tag/lock is a boundary-rule violation (see Stage 2 and Stage 5) and must be corrected.

## Stage 7 — ALARP

Apply automatically — **never leave ALARP blank**:
```
Residual LOW       → ALARP NO
Residual MEDIUM    → ALARP YES
Residual HIGH      → ALARP YES
Residual EXTREME   → do not approve the file; escalate immediately
```
Never mark YES on every row automatically, and never leave the cell empty for LOW — write the explicit NO. If Residual stays EXTREME: propose additional controls, re-score C/L, escalate if risk doesn't drop, and block final Excel issuance until resolved.

## Stage 7.5 — PPE (mandatory, own pass)

This is never skipped and never left implicit. Using `General_Instructions.md` (Part 3 — PPE Reference) and the project's established PPE-placement rule:

- **General/environmental PPE** (applies for the whole task/location — e.g. Safety Goggles, 3-Layer Mask or higher, Ear Plug, Safety Shoes, Safety Helmet, base CAT-2 clothing) is written **once**, as its own bullet list, inside the TBT (Toolbox Talk) step's Detail — as a PPE Briefing covering the full task duration.
- **Activity-specific PPE** (tied to a specific hazard/step — e.g. Arc-Rated Glove up to 500V Class-00 and CAT-2 Arc-Rated Clothing for any cable/terminal disconnect-connect step; Mechanical Glove + Safety Goggles + 3-Layer Mask for any rigging/lifting/manual-handling step; FBH + safe scaffold for work at height) is written **both** as an overview bullet in the TBT step **and** repeated in the `Detail` cell of every specific step where that activity actually occurs.
- Select the specific PPE items per the actual hazard/voltage/activity of each step (arc rating per the equipment's actual voltage, not a fixed default) — never a generic "wear PPE" bullet with no named items.
- A JSA with zero PPE bullets anywhere in the file, or with PPE only in a generic unnamed form, fails this stage — do not proceed to Excel build until every step needing PPE has it named.

## Stage 8 — Formatting

Organize these fields into short bullet points with `•`: Potential Hazard, Cause, Consequences, Control Measures Detail, Responsible. Split on commas/semicolons/independent items, but never split a single concept such as "Pinch/Crush Points" or "Dropped/Swinging Load". Keep every bullet short, clear, and non-repetitive. Any Job Step covering more than one discrete action/parameter uses a short colon-terminated title followed by its own bullet points — never a run-on comma/slash-joined sentence; a single atomic-action step stays one plain line.

## Stage 9 — Excel build (template fidelity)

**Mode governs the base file (see Stage -1):**
- **REVIEW mode:** build the approved corrections directly into the same accessible workbook — the chat attachment, the referenced/project file, or the workbook already open in an Excel Agent session — keeping its file identity, sheets, merges, formulas, and any content not flagged for correction untouched. Do not create a second, brand-new file from the blank `JSA Form.xlsx` template alongside or instead of the corrected accessible file.
- **CREATE mode:** copy the real blank `JSA Form.xlsx` (never open/edit the original template file itself) into the section's `Updated` subfolder and build the new content into that copy. This applies whenever the source material was a PDF, image, or plain description, exactly as it does for a fully spoken request — the deliverable is always an Excel file, never a text-only report, unless the user explicitly asks for a report instead.

`JPPMDES.xlsx` is read-only reference data in both modes — never the base file being built or corrected, and never itself delivered as the output, even if it happens to be the workbook currently open in an Excel Agent session.

In both modes, preserve: sheet names and order, merged cells, colors, borders, column widths/row heights, formulas, Risk Matrix, data-validation lists, print settings, header/footer. Never redesign the form. **Deliver in the way that matches the current surface:** in Claude Code/Claude Desktop, send the resulting file back into the conversation directly — never route it through Google Drive or any other external storage unless the user explicitly asks for that; in an Excel Agent/add-in session, save the corrections into the currently-open workbook in place rather than producing a separate chat file, unless the user asks for a separate file instead.

- Each Job Step occupies a 3-row block; never duplicate a Job Step per control measure; write all bullets for that step inside the block's merged cell.
- Do not assume merge ranges (e.g. Potential Hazard, Cause, Consequences, Detail, Responsible column spans) — inspect the actual template first, then replicate its structure literally.
- **Merge consistency across every column of a block (mandatory, every build — verify, never assume it "looks right"):** every column meant to span the whole Job Step (Job Step itself, Potential Hazard, Cause, Consequences, **Type**, Detail, Responsible — all of them, including Type, which is easy to miss since it's a short one-cell value rather than a bullet list) must be an actual Excel merge covering the exact same row range as that block's own boundary — same start row, same end row, block by block. A column that's left unmerged (or merged to a different, smaller/larger row range than its neighbors) will visually misalign even when its text content is correct, because its content then sits against the wrong sub-row once C/L/R's own two-row (Inherent/Residual) sub-split is laid over it. After writing a block, explicitly re-verify (don't assume from the copy step) that each of these columns' merge range matches the block's row range for that specific block — and do this as one pass across **every** block in the file, not spot-checked on a few. A block where one column shows content only in its first sub-row while its neighbors show content spread across all sub-rows is this defect, not a data problem — fix by merging that column to the full block range, not by re-typing its text.
- If new steps are needed, copy a correct block from the template (preserving formulas/formatting) rather than building one from scratch.
- **R-formula coverage (mandatory, every build):** the base template only pre-populates the R-risk formula (columns K and P) for the job-step blocks that existed when it was last updated, even though Data Validation and conditional formatting already span all 32 slots. For any newly-used block beyond that range, copy the exact formula string from a filled block (e.g. row 27) and substitute the row number into every cell reference, for both K (Inherent) and P (Residual). Confirm visually before delivery that every used step shows a colored LOW/MEDIUM/HIGH/EXTREME value, never a blank/grey cell.
- **C/L dropdown lists (mandatory, every build):** openpyxl silently drops the template's original x14 extended list data-validations on load/save. Before saving, explicitly re-add classic `DataValidation` (`type="list"`) objects: formula `'Lists'!$A$7:$A$11` applied to every Inherent/Residual **C** cell (columns I and N), and formula `'Lists'!$A$1:$A$5` applied to every Inherent/Residual **L** cell (columns J and O) — across all 32 job-step block anchor rows (27, 30, 33 … 120), matching the original template's scope, not just the filled rows. C and L values must never be left as hardcoded plain text with no dropdown.
- **Left alignment (mandatory, every cell written):** every text-bearing cell — Area/Equipment ID, Job Description, Job Location, Applicable Procedures, and each block's Potential Hazard, Cause, Consequences, Detail, and Responsible — must be set to horizontal-left alignment with `wrap_text=True`. Never leave these centered (the template's default), whether on first build or on a later edit.
- **Column width sanity check (mandatory, before computing row heights):** a too-narrow column inflates the wrapped-line count artificially — e.g. a 2-word control name like "PTW Verification" wrapping one word per line instead of fitting on one line — which then produces a needlessly tall row-height calculation downstream. Before doing the row-height pass below, check each bullet-bearing column's width (Potential Hazard, Cause, Consequences, Type, Detail, Responsible): a short, typical bullet (roughly 2-3 words) for that column must fit on one line, at most two. If it doesn't, widen that column — Detail in particular tends to be left too narrow relative to Consequences/Cause in the template — to roughly match the width of a comparable neighboring column, not squeezed. Do this once, for every such column, across the whole sheet, before the row-height calculation, since row heights computed against an unfixed narrow column will all need recomputing anyway.
- **Row-height auto-fit for merged cells (mandatory, every build — never fixed per-row manually):** `wrap_text=True` alone does not resize a merged cell's row height — Excel/openpyxl never auto-fits row height across a merge, so content clips visually even though it's fully present in the cell. For every Job Step's 3-row merged block, after all bullets are written **and column widths are sanity-checked (above)**: count the wrapped-line count each text column will actually take (bullets in that cell, plus any single bullet long enough to wrap within its column's width, given the column's character width) across Potential Hazard, Cause, Consequences, Detail, and Responsible; take the max line count among them; and set that block's total height (summed across its 3 physical rows, distributed evenly, e.g. ~15 points/line) to fit it. Apply this to **every** Job Step block in the file — not only the ones that look clipped on a quick visual check — since the same silent-clip risk exists in every merged block whenever the true line count exceeds the block's current row heights. This is a whole-file pass, done once after all content is written and before final QA, not a per-row manual fix.

## Stage 9.5 — Print & page layout (mandatory, own pass, after row heights are set in Stage 9)

Row heights set in Stage 9 (to fit wrapped bullet content) change how the sheet paginates — never leave print/page setup at whatever the template happened to have, and never fix this page-by-page by hand. Run this as one whole-file pass:

- **Row height mode:** Page Layout → Row Height must be **Automatic**, not a fixed/manual value carried over from the template or from a prior edit. A manually fixed row height is what causes a tall bullet block to get visually clipped again even after Stage 9 recalculated the content — Automatic lets Excel/the export engine size to content while still respecting the explicit heights Stage 9 set for merged blocks.
- **Hide unused rows:** hide every row below the last used Job Step block (and any genuinely empty row within range) before setting the print area — an unhidden trailing blank region prints extra blank pages and shifts the border/print-area logic below.
- **Thick outside border:** apply a thick (not thin/medium-default) outside border around the full data region actually in use — from the header row through the last used Job Step block's last row, across all used columns. Recompute this range after hiding rows, not before.
- **Print area = the thick-outside-border region:** set Print Area to exactly that bordered rectangle. Never leave Print Area stale from a previous version of the file when the used range has changed (fewer/more Job Steps, hidden rows).
- **No row may be split across a page break (mandatory — this is the most common visible defect):** Excel paginates purely by cumulative row height against the printable page height (paper size minus margins minus header/footer, divided by the scale factor) — it has no "keep block together" feature, so once Stage 9 gives blocks uneven heights, an automatic page break can fall inside a 3-row merged Job Step block, printing part of it on one page and the rest on the next. Prevent this on **every** page boundary in the file, not just the ones a spot-check happens to catch: compute the printable height per page, walk the rows in order accumulating height, and whenever the next Job Step block's rows would cross that boundary, insert an explicit horizontal page break **before** that block (pushing the whole block, all 3 of its rows, onto the next page) instead of letting Excel's automatic break fall wherever the cumulative height lands. Never split a Job Step's 3-row merged block across two pages under any circumstance.
- **Fit to width, automatic height:** Page Setup scaling = Fit to 1 page wide, height left unconstrained (not "Fit to 1 page tall", which would shrink text and reintroduce clipping) — combined with the row-height-automatic setting above.
- **Repeat header row:** set Print Titles so the column header row repeats at the top of every printed page — required whenever the file spans more than one page (e.g. a long Tasks Register export), so a page opened on its own is still readable.

This stage must be re-run whenever Stage 9 changes row heights or the used row range (new Job Steps added, rows hidden) — a print layout computed for a smaller/differently-sized file does not carry over correctly to a larger one.

## Stage 10 — Final QA (independent pass, after Excel is built)

Check all of:
1. The correct mode ran (REVIEW edited the accessible file — chat attachment, referenced file, or the open Excel Agent workbook — in place; CREATE built into a fresh copy of the blank template) — and this matches what was stated to the user in Stage -1
2. Every Job Step is complete
3. Sequence logic is correct and matches the Register's `Task Scope / Activities` order, and — whenever isolation is involved — the mandatory 10-step opening (Stage 2) is present in that exact order, with `Obtain Work Permit (PTW)` never first and TBT never before the permit
4. `Close PTW` exists and is last
5. Nothing appears after `Close PTW`
6. Isolation type is correct and matches the Register's `Isolation Required` field exactly — no extra isolation type added
7. C/L/R before and after control are correct and scaled to the actual job profile (Stage 4)
8. No Fatality/C5 in any lifting step
9. No automatic C5 for 480V work
10. ALARP is correct and never left blank
11. Roles are correct, only from the allowed list, and match the Register's `Lifting Equipment` field; Rigger/Signal Man appear only on steps that actually use a **Mobile Crane** — never on a Forklift-only step, and never on an Overhead Crane step either
12. No disallowed roles anywhere
13. Bullets present and concise
14. No invented data; no "TBC"/placeholder text for any fact the Register already states
15. No unresolved EXTREME residual risk
16. Excel matches the original template (or, in REVIEW mode, the accessible file's own original structure)
17. No clipped cells or hidden text — every merged Job Step block's row height was explicitly set to fit its actual max line count (Stage 9), not left at the template's default/unadjusted height
18. No formula errors (#REF!, #VALUE!, #DIV/0!)
19. Displayed values match the Risk Matrix
20. C/J/N/O cells are dropdown data-validation lists, not hardcoded plain text
21. Text-bearing cells (headers and Hazard/Cause/Consequences/Detail/Responsible) are left-aligned, not centered
22. Job Description, Area/Equipment ID, and Job Location are copied verbatim from the Register row, unchanged in wording (or are correctly left empty/flagged rather than invented, only when no Register row matches)
23. Every used Job Step block (including any beyond the template's originally pre-filled range) has a working, colored R value in both K and P — never blank/grey
24. Zero mention of coupling, alignment, or the Mechanical department anywhere in the JSA (Job Step, Hazard, Cause, Consequence, Detail, or Responsible) for any equipment group except 88TK/88PF — not even a hand-off or confirmation step
25. A separate rotation-test step exists for Electrical's own scope, placed after Electrical's reconnection and before final inspection, whenever Electrical is responsible for it
26. PPE is present: a PPE Briefing bullet list in the TBT step, plus activity-specific PPE repeated in every step where that activity occurs
27. The delivered file matches its surface: sent directly into the conversation (not routed through Google Drive or other external storage) in Claude Code/Claude Desktop, or saved in place into the currently-open workbook in an Excel Agent session
28. No step has the Permit Receiver "applying" the boundary tag or isolation/red lock — that action is Issuer-only; the Receiver's step reads as `Lock/Tag Verification` then `Personal Lock`
29. An Excel deliverable was produced even though the source material was a non-Excel document (PDF/image/description) — no report-only fallback and no question asked about whether Excel output is possible
30. `JPPMDES.xlsx` was used only as a read-only reference source — never edited, corrected, or delivered as the output workbook
31. Print/page layout (Stage 9.5) was run after the final row heights and used-row range were set: Row Height is Automatic (not manually fixed), unused rows are hidden, a thick outside border and Print Area cover exactly the used data region, Fit-to-width/automatic-height scaling is set, header row repeats on every page, and — checked page by page, not just spot-checked — no Job Step's 3-row merged block is split across a page break
32. Merge ranges are consistent across every column of every Job Step block, checked block by block for the whole file: Job Step, Potential Hazard, Cause, Consequences, **Type**, Detail, and Responsible each merge across the exact same row range as that block — no column left unmerged or merged to a mismatched range while its neighbors are correct (Type is checked explicitly, not just the bullet-list columns)
33. No bullet-bearing column (Potential Hazard, Cause, Consequences, Type, Detail, Responsible) is narrow enough that a short 2-3 word bullet wraps one word per line — checked whole-file, not spot-checked; Detail is the one most often left too narrow relative to Consequences/Cause

Escalate and refuse to approve if any of: the wrong mode ran (a review request produced a brand-new file instead of an edited accessible file, or vice versa); weight/data conflict between Register/PDF/Excel/user; unconfirmed equipment number, voltage, or weight when the Register doesn't resolve it; C5 in a lifting step; a disallowed role; a role not matching the Register's lifting equipment; `Close PTW` not last; a test after `Close PTW`; Residual EXTREME; an unjustified C or L reduction or an unjustified inflated C/L not matching the job's actual profile; a material template deviation; a formula failure; inability to verify source content; missing C/L dropdown validation; centered text-bearing cells; a Job Description/Area/Location that deviates in wording from an existing Register row; a blank ALARP cell; a missing R value for any used step; an isolation type not in the Register; any mention of coupling/alignment/Mechanical outside the 88TK/88PF exception, including a hand-off step; a missing rotation-test step for Electrical's own scope; missing or generic-only PPE; delivery that doesn't match the surface (routed through Google Drive/external storage instead of delivered in the conversation on chat surfaces, or built as a separate new file instead of saved into the open workbook on an Excel Agent surface without the user asking for a separate file); the Permit Receiver written as applying the Issuer's boundary tag/isolation lock instead of verifying it; a non-Excel source document treated as a reason to skip or question the Excel deliverable; `JPPMDES.xlsx` treated as the JSA workbook under review/creation instead of as reference data; a Job Step block split across a page break; Row Height left manually fixed instead of Automatic; Print Area not matching the actual used/bordered region; unused rows left visible/unhidden; no repeating header row on a multi-page file; a block where one column's merge range doesn't match its neighbors' (unmerged or mismatched row span for Job Step/Potential Hazard/Cause/Consequences/Type/Detail/Responsible — Type included); a bullet-bearing column left narrow enough that short bullets wrap one word per line.

**Pre-delivery self-test — these must all FAIL as expected:**
```
Accessible filled file (attached/referenced/open in Excel Agent) answered by building a brand-new file from blank template → FAIL
Delivered via Google Drive/external storage instead of in the conversation (chat surfaces) → FAIL
Excel Agent open-workbook review delivered as a separate new file instead of saved in place → FAIL
Asking "how do I deliver this" because the source was a PDF, not an Excel       → FAIL
"Obtain Work Permit (PTW)" as the first Job Step                → FAIL
Toolbox Talk written before the Work Permit is obtained          → FAIL
Lifting step + Fatality                                     → FAIL
480V isolation + automatic C5                                → FAIL
Residual LOW + ALARP YES                                     → FAIL
Residual LOW + ALARP blank                                   → FAIL
Close PTW + test-run in same step                            → FAIL
Any step after Close PTW                                     → FAIL
Job Supervisor as Responsible                                → FAIL
Banksman as Responsible                                       → FAIL
Fixed weight without confirmation                             → FAIL
Long narrative in Detail                                       → FAIL
Rebuilt template instead of original template                → FAIL
C/L cell as hardcoded text (no dropdown)                       → FAIL
Centered Hazard/Cause/Consequences/Detail/Responsible cell     → FAIL
Job Description reworded instead of copied from Register row   → FAIL
Run-on comma/slash Job Step sentence instead of bullets        → FAIL
Used step beyond original range with blank/grey R cell         → FAIL
"TBC"/placeholder for a fact the Register already states       → FAIL
Mechanical isolation added when Register says electrical-only  → FAIL
Any "Confirm Mechanical..." hand-off step (non-88TK/88PF)       → FAIL
Heavy C4/C5 residual across a light forklift-only task          → FAIL
Missing rotation-test step for Electrical's own scope            → FAIL
Zero PPE bullets anywhere in the file                            → FAIL
Permit Receiver "applying" boundary tag / isolation red lock     → FAIL
JPPMDES.xlsx treated as the JSA file to review/create/deliver    → FAIL
```

## What to show the user

Do not dump the full internal per-stage output unless explicitly asked. Deliver only:
1. The mode that ran (REVIEW of the accessible file — chat attachment, referenced file, or the open Excel Agent workbook — or CREATE of a new one) and a short summary of what was checked, including which source-of-truth files were found and used (Stage 0)
2. The list of unconfirmed data
3. The list of material corrections made
4. QA result: **PASS** or **BLOCKED**
5. The final Excel file, delivered the way that matches the surface: sent directly in the conversation (chat surfaces) or saved in place into the open workbook (Excel Agent) — the same accessible file, corrected, in REVIEW mode; a new file built from the blank template in CREATE mode
