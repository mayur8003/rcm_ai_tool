"""
RCM Redrafting & Design Review Tool — Groq API version
Run with: streamlit run rcm_tool_groq.py

Install dependencies:
    pip install streamlit pandas openpyxl groq pydantic reportlab python-docx
"""

import io
import json
import time
import pandas as pd
import streamlit as st
from groq import Groq
from pydantic import BaseModel, Field

# ---------------------------------------------------------------------------
# Page config
# ---------------------------------------------------------------------------
st.set_page_config(page_title="RCM AI Tool", layout="wide")
st.title("📊 Risk & Control Matrix (RCM) AI Tool")
st.caption("Powered by Groq AI  |  Big 4 Audit Methodology  |  ICFR / SOX / IFC")

# ---------------------------------------------------------------------------
# Sidebar — Settings
# ---------------------------------------------------------------------------
st.sidebar.header("⚙️ Settings")

# Auto-read from Streamlit Cloud secrets if configured
_secret_key = st.secrets.get("GROQ_API_KEY", "") if hasattr(st, "secrets") else ""

api_key = st.sidebar.text_input(
    "Groq API Key",
    value=_secret_key,
    type="password",
    help=(
        "Get your free key from https://console.groq.com/keys  |  "
        "Or store it permanently via Streamlit Cloud → Settings → Secrets"
    )
)

MODEL_OPTIONS = {
    "Llama 3.3 70B (Best Quality)": "llama-3.3-70b-versatile",
    "Llama 3.1 8B (Fastest)":       "llama-3.1-8b-instant",
    "Mixtral 8x7B (Balanced)":      "mixtral-8x7b-32768",
    "Gemma 2 9B":                   "gemma2-9b-it",
}
selected_label = st.sidebar.selectbox("Model", list(MODEL_OPTIONS.keys()))
selected_model = MODEL_OPTIONS[selected_label]

batch_size = st.sidebar.slider(
    "Rows per batch", 5, 30, 10, step=5,
    help="Smaller = safer for rate limits"
)
retry_delay = st.sidebar.slider("Retry delay (seconds)", 5, 60, 15, step=5)
max_retries = st.sidebar.number_input("Max retries per batch", 1, 5, 3)

# ---------------------------------------------------------------------------
# Sidebar — Tool Mode
# ---------------------------------------------------------------------------
st.sidebar.markdown("---")
st.sidebar.header("🛠️ Tool Mode")
tool_mode = st.sidebar.radio(
    "Select what you want to do:",
    [
        "✏️  Redraft & Classify Controls",
        "🔍  Full RCM Design Effectiveness Review",
    ],
    help=(
        "Redraft: rewrites every control row in Big 4 style and classifies "
        "Key/Non-Key and Anti-Fraud.\n\n"
        "Design Review: performs a 14-part ICFR/SOX design effectiveness review "
        "and produces a professional audit report."
    ),
)

# ---------------------------------------------------------------------------
# Sidebar — Process Context (used by both modes, mandatory for Design Review)
# ---------------------------------------------------------------------------
st.sidebar.markdown("---")
st.sidebar.header("📋 Process Context")

PROCESS_OPTIONS = [
    "Purchase to Pay (P2P)",
    "Order to Cash (O2C)",
    "Record to Report (R2R)",
    "Hire to Retire (H2R)",
    "Inventory & Warehouse Management",
    "Fixed Assets",
    "Treasury & Cash Management",
    "Tax & Statutory Compliance",
    "Financial Reporting & Close",
    "Payroll",
    "Revenue Recognition",
    "Borrowings & Debt Management",
    "Capital Expenditure (Capex)",
    "Other (specify below)",
]

INDUSTRY_OPTIONS = [
    "Steel Manufacturing",
    "FMCG",
    "Pharmaceutical",
    "Automobile / Auto Components",
    "IT / Software Services",
    "NBFC / Financial Services",
    "Real Estate & Construction",
    "Power & Energy",
    "Retail",
    "Cement & Infrastructure",
    "Textile",
    "Chemicals",
    "Other (specify below)",
]

ERP_OPTIONS = [
    "SME Assist",
    "SAP S/4HANA",
    "SAP ECC",
    "Oracle Fusion",
    "Oracle EBS",
    "Microsoft Dynamics 365",
    "Tally Prime",
    "Zoho Books",
    "QuickBooks",
    "Custom / In-house ERP",
    "No ERP / Manual",
    "Other (specify below)",
]

REVIEW_PART_OPTIONS = [
    "All Parts (Full Report)",
    "Part 1  – Process Understanding",
    "Part 2  – Risk Review",
    "Part 3  – Control Review",
    "Part 4  – Missing Controls",
    "Part 5  – Unnecessary Controls",
    "Part 6  – Key vs Non-Key Classification",
    "Part 7  – Anti-Fraud Review",
    "Part 8  – Segregation of Duties",
    "Part 9  – Control Redrafting",
    "Part 10 – Financial Statement Assertions",
    "Part 11 – COSO Mapping",
    "Part 12 – Control Testing Procedures",
    "Part 13 – Automation Opportunities",
    "Part 14 – Final Executive Report",
]

selected_process = st.sidebar.selectbox("Business Process", PROCESS_OPTIONS)
custom_process   = ""
if selected_process == "Other (specify below)":
    custom_process = st.sidebar.text_input("Specify Process")

selected_industry = st.sidebar.selectbox("Industry", INDUSTRY_OPTIONS)
custom_industry   = ""
if selected_industry == "Other (specify below)":
    custom_industry = st.sidebar.text_input("Specify Industry")

selected_erp  = st.sidebar.selectbox("ERP / System", ERP_OPTIONS)
custom_erp    = ""
if selected_erp == "Other (specify below)":
    custom_erp = st.sidebar.text_input("Specify ERP / System")

# For Design Review: which parts to generate
selected_parts: list[str] = []
if "Design" in tool_mode:
    st.sidebar.markdown("---")
    st.sidebar.header("📑 Review Scope")
    selected_parts = st.sidebar.multiselect(
        "Select report parts to generate:",
        REVIEW_PART_OPTIONS,
        default=["All Parts (Full Report)"],
        help="Select one or more parts. Choose 'All Parts' for the complete report.",
    )
    if not selected_parts:
        selected_parts = ["All Parts (Full Report)"]

# Resolve final values
process_label  = custom_process  if selected_process  == "Other (specify below)" else selected_process
industry_label = custom_industry if selected_industry == "Other (specify below)" else selected_industry
erp_label      = custom_erp      if selected_erp      == "Other (specify below)" else selected_erp

# ---------------------------------------------------------------------------
# Pydantic schema  (Redraft mode)
# ---------------------------------------------------------------------------
class ControlItem(BaseModel):
    original_id: int
    redrafted_control: str = Field(
        description="Audit-ready control description specifying Who, What, When, and How."
    )
    anti_fraud: str = Field(
        description="'Y' if control prevents/detects fraud/misappropriation/override, else 'N'."
    )
    risk_rating: str = Field(
        description="'High', 'Medium', or 'Low' based on potential impact."
    )
    control_type: str = Field(
        description="'Key' or 'Non-Key' per ICFR/SOX classification criteria."
    )

class RCMBatchResponse(BaseModel):
    items: list[ControlItem]

def apply_control_type_rule(anti_fraud: str, risk_rating: str, ai_control_type: str) -> str:
    return ai_control_type

# ---------------------------------------------------------------------------
# Prompt constants
# ---------------------------------------------------------------------------
BIG4_REDRAFTING_PROMPT = """
=== BIG 4 AUDIT METHODOLOGY — CONTROL REDRAFTING RULES ===

You are an Internal Financial Controls (IFC)/ICFR/RCM specialist.
Rewrite each control description using Big 4 audit methodology.

The redrafted control MUST clearly answer ALL of the following:
• WHO performs the control?
• WHAT is performed?
• HOW is it performed?
• WHAT is reviewed?
• WHO reviews/approves it?
• WHEN / Frequency (if mentioned in the original)?
• WHAT evidence is generated?
• WHAT risk does the control mitigate?

Use the following writing style:
"The <Control Owner> prepares/verifies/reconciles/reviews <document/report/calculation>.
The <Reviewer/Approver> reviews and approves the same after verifying <specific checks>.
Upon approval, the transaction is processed/recorded/filed in <ERP/System>.
This ensures <control objective>."

Mandatory rules:
1. Use formal ICFR language throughout.
2. Do NOT use passive voice.
3. No grammatical errors.
4. Mention ERP or SME Assist (or the relevant system) wherever applicable.
5. Mention statutory compliance (Income Tax, GST, TDS, Companies Act, etc.) wherever applicable.
6. Keep the meaning exactly the same as the original — do not change the intent of the control.
7. Do NOT omit approvals mentioned or implied in the original.
8. Do NOT add assumptions or invent details not present in the original.
=== END OF REDRAFTING RULES ===
"""

ANTI_FRAUD_CRITERIA = """
=== ICFR/SOX ANTI-FRAUD CONTROL CLASSIFICATION CRITERIA ===

You are an ICFR/SOX expert.
For each control, determine whether it is an Anti-Fraud Control.

A control IS Anti-Fraud (Y) if it directly prevents or detects ANY of the following:
- Unauthorized transactions
- Management override of controls
- Fraudulent financial reporting
- Misappropriation of assets
- Fake vendors, customers, or employees
- Unauthorized payments
- Unauthorized changes to master data (vendor, customer, employee, bank details)
- Duplicate or fictitious payments
- Journal entry manipulation or unauthorized journal entries

A control is NOT Anti-Fraud (N) if it ONLY does one or more of the following
and does NOT address any of the fraud indicators above:
- Ensures mathematical accuracy or arithmetic correctness
- Ensures statutory or regulatory compliance (GST, TDS, Income Tax, Companies Act, etc.)
- Performs routine reconciliations without any fraud detection element
- Reviews reports purely for completeness or timeliness
- Supports operational efficiency or process adherence

CRITICAL RULES:
1. Assess each control independently — do NOT default all to Y or all to N.
2. If a control performs both a routine function AND has a fraud-prevention element
   (e.g., a reconciliation that also detects unauthorized entries), classify as Y.
3. A management review of journal entries, payments, or master data changes is
   Anti-Fraud (Y) because it can detect management override or unauthorized transactions.
4. A purely mathematical check or a report generation step with no fraud-detection
   element is NOT Anti-Fraud (N).
=== END OF ANTI-FRAUD CRITERIA ===
"""

ICFR_SOX_CRITERIA = """
=== ICFR / SOX KEY CONTROL CLASSIFICATION CRITERIA ===

You are an ICFR/SOX Risk and Control Matrix expert.
Classify each control as either Key Control or Non-Key Control using the criteria below.

A control is a KEY CONTROL if ANY of the following apply:
1. It prevents or detects a material misstatement in financial reporting.
2. It addresses a significant financial reporting risk.
3. It involves management review or approval.
4. It relates to statutory or regulatory compliance (Income Tax, GST, TDS, Companies Act,
   or equivalent regulations) where failure may result in material penalties or misstatement.
5. It is the primary control over a significant account balance or financial statement assertion.
6. There is no other compensating control that would detect the error if this control fails.

A control is a NON-KEY CONTROL if ALL of the following apply:
1. It is only an operational or administrative activity.
2. It is only data preparation or report generation with no direct financial reporting impact.
3. It is a supporting activity AND another review control already exists that mitigates the risk.
4. Failure would NOT reasonably result in a material financial reporting error.

CRITICAL RULE — NEVER classify as Non-Key:
- Management review controls over tax computations, financial statements, journal entries,
  provisions, reconciliations, bank payments, statutory returns, or significant estimates
  UNLESS there is a separate higher-level review control that directly mitigates the same risk.

Apply these criteria independently for each control. Do not default all controls to Key or
Non-Key — assess each one on its own merits.
=== END OF CLASSIFICATION CRITERIA ===
"""

def build_design_review_prompt(
    process: str,
    industry: str,
    erp: str,
    parts: list[str],
    rcm_text: str,
) -> str:
    """Build the full 14-part RCM Design Effectiveness Review prompt."""

    # If "All Parts" selected, include every part
    all_parts = "All Parts (Full Report)" in parts
    def include(part_label: str) -> bool:
        return all_parts or any(part_label in p for p in parts)

    scope_note = (
        "Generate ALL 14 parts of the Design Effectiveness Review."
        if all_parts
        else "Generate ONLY the following parts: " + ", ".join(parts) + "."
    )

    parts_block = ""

    if include("Part 1"):
        parts_block += """
──────────────────────────────
PART 1 – PROCESS UNDERSTANDING
──────────────────────────────
Identify:
• Business Process
• Process Objective
• Major Financial Statement Accounts impacted
• Significant Risks
• Process Boundaries
• Upstream and Downstream Processes
"""

    if include("Part 2"):
        parts_block += """
──────────────────────────────
PART 2 – RISK REVIEW
──────────────────────────────
Identify:
• Missing Risks
• Duplicate Risks
• Incorrectly Defined Risks
• Risks without Controls
• Risks that are over-controlled
• Risks not linked to Financial Reporting
Give reasons for every observation.
"""

    if include("Part 3"):
        parts_block += """
──────────────────────────────
PART 3 – CONTROL REVIEW
──────────────────────────────
For every control evaluate:
✓ Does the control actually mitigate the risk?
✓ Is the control preventive or detective?
✓ Is the control manual, automated, or IT-dependent?
✓ Is the control complete?
✓ Is the control written in professional RCM language?
✓ Is the control measurable and testable?
✓ Is the control missing any review step?
✓ Is evidence generated?
✓ Is the control owner appropriate?
✓ Is the reviewer/approver appropriate?
✓ Is the approval hierarchy sufficient?
"""

    if include("Part 4"):
        parts_block += """
──────────────────────────────
PART 4 – IDENTIFY MISSING CONTROLS
──────────────────────────────
Compare the RCM against industry best practices.
Identify missing controls related to:
Approvals, Authorizations, Reconciliations, Maker-Checker, ERP validations,
System controls, Management Reviews, Statutory Compliance, Monitoring Controls,
Exception Reports, Master Data Controls, Journal Entries, Bank Controls, Tax Controls,
Inventory Controls, Borrowings, Financial Closing, Provisions, Credit Limits,
Customer Master, Vendor Master, Payroll, Fixed Assets, Revenue Recognition, Fraud Prevention.

For every missing control provide a table with columns:
Risk | Suggested Control Description | Control Owner | Reviewer | Frequency |
Preventive/Detective | Manual/Automated | Evidence | Key/Non-Key | Anti-Fraud |
Financial Assertion | COSO Component
"""

    if include("Part 5"):
        parts_block += """
──────────────────────────────
PART 5 – IDENTIFY UNNECESSARY CONTROLS
──────────────────────────────
Identify controls which are:
• Operational only
• Documentation only
• Duplicate
• Administrative
• Do not reduce financial reporting risk
Explain why each control is unnecessary.
"""

    if include("Part 6"):
        parts_block += """
──────────────────────────────
PART 6 – KEY vs NON-KEY CLASSIFICATION
──────────────────────────────
Classify every control as Key or Non-Key.

KEY CONTROL if it:
• Prevents or detects material misstatement
• Mitigates significant financial reporting risk
• Has management review
• Relates to statutory compliance
• No compensating control exists
• Covers a significant account balance
• Serves a fraud prevention purpose

NON-KEY if it is:
• Administrative / Operational / Supporting / Documentation only
• Low financial reporting impact

Provide reason for every classification.
NEVER classify a management review of journal entries, payments, tax computations,
provisions, or reconciliations as Non-Key unless a higher-level compensating control exists.
"""

    if include("Part 7"):
        parts_block += """
──────────────────────────────
PART 7 – ANTI-FRAUD REVIEW
──────────────────────────────
Identify:
• Existing Anti-Fraud Controls
• Missing Anti-Fraud Controls
• Fraud Risks present in the process
• Management Override Risks
• Segregation of Duties Issues
• Fake Vendor / Fake Customer Risk
• Unauthorized Payments / Duplicate Payments
• Manual Journal Risk
• Revenue Manipulation
• Inventory Theft / Cash Misappropriation
"""

    if include("Part 8"):
        parts_block += """
──────────────────────────────
PART 8 – SEGREGATION OF DUTIES
──────────────────────────────
Check whether the same person performs any of:
• Preparation & Approval
• Vendor Creation & Payment
• Customer Creation & Credit Approval
• Bank Beneficiary Addition & Payment Release
• Journal Entry Preparation & Approval
• Inventory Custody & Accounting
• Asset Custody & Accounting
• Payroll Preparation & Approval
Highlight all SoD conflicts with recommended remediation.
"""

    if include("Part 9"):
        parts_block += """
──────────────────────────────
PART 9 – CONTROL REDRAFTING
──────────────────────────────
For every weak or incomplete control, rewrite it in Big 4 style.
Each redrafted control must clearly state:
WHO | WHAT | HOW | REVIEW | APPROVAL | SYSTEM | EVIDENCE | CONTROL OBJECTIVE
Use formal ICFR language and active voice. Do not change the business intent.
"""

    if include("Part 10"):
        parts_block += """
──────────────────────────────
PART 10 – FINANCIAL STATEMENT ASSERTIONS
──────────────────────────────
For every control identify which assertions it covers:
• Existence
• Completeness
• Accuracy
• Valuation
• Rights & Obligations
• Presentation & Disclosure
"""

    if include("Part 11"):
        parts_block += """
──────────────────────────────
PART 11 – COSO MAPPING
──────────────────────────────
Map every control to the COSO 2013 Framework component:
• Control Environment
• Risk Assessment
• Control Activities
• Information & Communication
• Monitoring Activities
"""

    if include("Part 12"):
        parts_block += """
──────────────────────────────
PART 12 – CONTROL TESTING PROCEDURES
──────────────────────────────
For every Key Control provide:
• Testing Procedure (step-by-step)
• Audit Evidence required
• Sample Size / Requirement
• Expected Result
• Possible Exceptions to watch for
"""

    if include("Part 13"):
        parts_block += """
──────────────────────────────
PART 13 – AUTOMATION OPPORTUNITIES
──────────────────────────────
Identify controls that can be automated using {erp} / ERP system controls.
For each, suggest the specific system configuration, workflow rule, or automated
control that would replace or strengthen the manual control.
""".format(erp=erp)

    if include("Part 14"):
        parts_block += """
──────────────────────────────
PART 14 – FINAL EXECUTIVE REPORT
──────────────────────────────
Prepare a professional executive report containing:
1.  Executive Summary
2.  Missing Controls Summary
3.  Weak Controls Summary
4.  Duplicate Controls
5.  Unnecessary Controls
6.  SoD Issues
7.  Anti-Fraud Gaps
8.  Key Control Summary
9.  Process Maturity Score (1–5 scale with justification)
10. Overall Design Effectiveness Rating (Effective / Partially Effective / Ineffective)
11. Top 10 Recommendations (ranked by priority)
12. Priority Matrix:
    • Critical
    • High
    • Medium
    • Low

Present all observations in professional tables suitable for an Internal Audit / ICFR report.
Maintain a Big 4 consulting style throughout.
"""

    prompt = f"""You are a Senior Manager in Risk Advisory (ICFR / IFC / SOX) with 20+ years of
experience at Big 4 firms (EY, PwC, Deloitte, KPMG). You have expertise in designing, reviewing,
and testing Risk and Control Matrices (RCMs) for manufacturing, steel, FMCG, NBFC, and service
industries.

ENGAGEMENT CONTEXT:
• Business Process : {process}
• Industry         : {industry}
• ERP / System     : {erp}

BENCHMARK FRAMEWORKS:
• COSO 2013 Framework
• Internal Financial Controls (IFC) under the Companies Act, 2013
• ICAI Guidance Note on IFC
• SOX Best Practices
• Standard Industry Practices for {industry}
• Financial Statement Assertions
• Anti-Fraud Control Framework

TASK:
{scope_note}
Perform a comprehensive Design Effectiveness Review of the RCM provided below.
Present all findings in professional tables with clear headings, suitable for inclusion
in an Internal Audit or ICFR report. Maintain a Big 4 consulting style throughout.

{parts_block}

──────────────────────────────
RCM DATA FOR REVIEW:
──────────────────────────────
{rcm_text}
"""
    return prompt


# ---------------------------------------------------------------------------
# Core: call Groq — Redraft mode (structured JSON)
# ---------------------------------------------------------------------------
def call_groq(
    client: Groq,
    model: str,
    controls_input: list[dict],
    max_retries: int,
    retry_delay: int,
    has_risk_col: bool = False,
    process: str = "",
    industry: str = "",
    erp: str = "",
) -> RCMBatchResponse:

    risk_instruction = (
        """The 'existing_risk' field contains the pre-defined risk description for each control.
Use it to:
- Better understand the nature and severity of the risk being mitigated
- Align the Risk Rating with the defined risk (e.g. if risk mentions financial loss → High)
- Improve Anti-Fraud assessment based on whether the risk involves fraud, misappropriation, or override
- Inform the Key / Non-Key classification where the risk description clarifies materiality
"""
        if has_risk_col else ""
    )

    context_block = ""
    if process or industry or erp:
        context_block = f"""
ENGAGEMENT CONTEXT:
• Business Process : {process or 'Not specified'}
• Industry         : {industry or 'Not specified'}
• ERP / System     : {erp or 'Not specified'}
Use this context to improve the accuracy of ERP references, industry-specific language,
and statutory compliance mentions in the redrafted controls.
"""

    prompt = f"""You are a senior internal audit expert with deep knowledge of fraud risk assessment
and ICFR/SOX internal controls over financial reporting.

Review and redraft the following Risk and Control Matrix (RCM) control activities.

{context_block}
{risk_instruction}
{BIG4_REDRAFTING_PROMPT}
{ANTI_FRAUD_CRITERIA}
{ICFR_SOX_CRITERIA}

For each control, independently assess and return:

1. REDRAFTED CONTROL: Rewrite using the Big 4 audit methodology and writing style defined above.
   The redrafted control MUST answer: WHO, WHAT, HOW, WHAT is reviewed, WHO approves,
   WHEN/Frequency, WHAT evidence is generated, and WHAT risk is mitigated.
   Follow the sentence structure:
   "The <Control Owner> prepares/verifies/reconciles/reviews <document/report/calculation>.
   The <Reviewer/Approver> reviews and approves the same after verifying <specific checks>.
   Upon approval, the transaction is processed/recorded/filed in <ERP/System>.
   This ensures <control objective>."
   Use formal ICFR language, active voice, no assumptions, no omissions.
   Reference the ERP ({erp or 'the ERP system'}) wherever applicable.

2. ANTI-FRAUD (Y or N): Apply EXACTLY the Anti-Fraud classification criteria defined above.
   Mark 'Y' if the control prevents or detects any of the fraud indicators listed above
   (unauthorized transactions, management override, fictitious vendors/payments, journal
   entry manipulation, misappropriation of assets, unauthorized master data changes, etc.).
   Mark 'N' if the control ONLY ensures mathematical accuracy, statutory compliance,
   routine reconciliation, report completeness, or operational efficiency — with no
   fraud-detection element whatsoever.
   If a control serves both a routine purpose AND contains a fraud-detection element, mark 'Y'.
   DO NOT default all to Y or all to N — assess each control independently.

3. RISK RATING (High, Medium, or Low):
   - High: Material financial impact or regulatory breach possible
   - Medium: Moderate impact, recoverable with effort
   - Low: Minor impact, easily corrected
   If 'existing_risk' is provided, use it to inform your rating.

4. CONTROL TYPE (Key or Non-Key):
   Apply EXACTLY the ICFR/SOX classification criteria stated above.
   - Key: Meets one or more of the Key Control criteria listed above.
   - Non-Key: Meets ALL of the Non-Key criteria listed above.
   Remember: NEVER classify a management review control over tax computations, financial
   statements, journal entries, provisions, reconciliations, bank payments, statutory returns,
   or significant estimates as Non-Key unless a separate higher-level review directly mitigates
   the same risk.

Preserve original_id exactly as given.

Return ONLY a valid JSON object in this exact format (no markdown, no explanation):
{{
  "items": [
    {{
      "original_id": <int>,
      "redrafted_control": "<string>",
      "anti_fraud": "<Y or N>",
      "risk_rating": "<High or Medium or Low>",
      "control_type": "<Key or Non-Key>"
    }}
  ]
}}

Controls to process:
{json.dumps(controls_input, indent=2)}"""

    last_error = None
    for attempt in range(1, max_retries + 1):
        try:
            response = client.chat.completions.create(
                model=model,
                messages=[
                    {
                        "role": "system",
                        "content": (
                            "You are a senior internal audit expert specialising in ICFR/SOX "
                            "internal controls over financial reporting, trained in Big 4 audit "
                            "methodology for drafting Risk and Control Matrix (RCM) controls. "
                            "When redrafting controls, always use formal ICFR language, active "
                            "voice, and the Big 4 writing style: WHO performs, WHAT is done, "
                            "HOW it is done, WHO reviews/approves, WHEN, WHAT evidence is "
                            "generated, and WHAT risk is mitigated. "
                            "For Anti-Fraud classification, apply the defined criteria strictly: "
                            "Y only if the control prevents or detects unauthorized transactions, "
                            "management override, fictitious vendors/payments, journal entry "
                            "manipulation, or misappropriation of assets; N if it only ensures "
                            "accuracy, compliance, or operational efficiency. Never default all "
                            "controls to Y or N — assess each one independently. "
                            "Apply the ICFR/SOX Key Control classification criteria strictly and "
                            "independently for each control item. "
                            "Always respond with valid JSON only."
                        )
                    },
                    {"role": "user", "content": prompt}
                ],
                temperature=0.2,
                response_format={"type": "json_object"},
            )

            raw = response.choices[0].message.content.strip()
            if raw.startswith("```"):
                raw = raw.split("```")[1]
                if raw.startswith("json"):
                    raw = raw[4:]
            raw = raw.strip()

            data = json.loads(raw)
            return RCMBatchResponse.model_validate(data)

        except Exception as e:
            last_error = e
            err = str(e)
            if "429" in err or "rate_limit" in err.lower():
                wait = retry_delay * attempt
                st.warning(f"⏳ Rate limit hit (attempt {attempt}/{max_retries}). Waiting {wait}s…")
                time.sleep(wait)
            else:
                raise e

    raise Exception(f"All {max_retries} retries exhausted.\nLast error: {last_error}")


# ---------------------------------------------------------------------------
# Core: call Groq — Design Review mode (free-text streaming)
# ---------------------------------------------------------------------------
def call_groq_review(
    client: Groq,
    model: str,
    prompt: str,
    max_retries: int,
    retry_delay: int,
) -> str:
    last_error = None
    for attempt in range(1, max_retries + 1):
        try:
            response = client.chat.completions.create(
                model=model,
                messages=[
                    {
                        "role": "system",
                        "content": (
                            "You are a Senior Manager in Risk Advisory (ICFR/IFC/SOX) with 20+ "
                            "years of experience at Big 4 firms (EY, PwC, Deloitte, KPMG). "
                            "You produce comprehensive, professional RCM Design Effectiveness "
                            "Review reports using Big 4 consulting standards. "
                            "Always use structured tables, clear headings, and formal audit language. "
                            "Never use casual language. Be thorough and specific."
                        )
                    },
                    {"role": "user", "content": prompt}
                ],
                temperature=0.3,
                max_tokens=8000,
            )
            return response.choices[0].message.content.strip()

        except Exception as e:
            last_error = e
            err = str(e)
            if "429" in err or "rate_limit" in err.lower():
                wait = retry_delay * attempt
                st.warning(f"⏳ Rate limit hit (attempt {attempt}/{max_retries}). Waiting {wait}s…")
                time.sleep(wait)
            else:
                raise e

    raise Exception(f"All {max_retries} retries exhausted.\nLast error: {last_error}")


# ---------------------------------------------------------------------------
# Export helpers — PDF (reportlab) and Word (python-docx)
# ---------------------------------------------------------------------------
def _parse_md_lines(md_text: str):
    """
    Parse Markdown text into a list of (type, content) tuples.
    Types: h1, h2, h3, hr, table_row, bullet, bold_line, normal
    """
    lines = md_text.splitlines()
    parsed = []
    i = 0
    while i < len(lines):
        line = lines[i]
        stripped = line.strip()
        if not stripped:
            i += 1
            continue
        if stripped.startswith("### "):
            parsed.append(("h3", stripped[4:].strip()))
        elif stripped.startswith("## "):
            parsed.append(("h2", stripped[3:].strip()))
        elif stripped.startswith("# "):
            parsed.append(("h1", stripped[2:].strip()))
        elif stripped.startswith("---") and len(stripped) >= 3 and all(c == "-" for c in stripped):
            parsed.append(("hr", ""))
        elif stripped.startswith("|"):
            # table row
            cells = [c.strip() for c in stripped.strip("|").split("|")]
            # skip separator rows like |---|---|
            if not all(set(c.replace("-","").replace(":","").replace(" ","")) == set() for c in cells):
                parsed.append(("table_row", cells))
        elif stripped.startswith("- ") or stripped.startswith("* "):
            parsed.append(("bullet", stripped[2:].strip()))
        elif stripped.startswith("**") and stripped.endswith("**") and len(stripped) > 4:
            parsed.append(("bold_line", stripped.strip("*")))
        else:
            parsed.append(("normal", stripped))
        i += 1
    return parsed


def generate_pdf(md_text: str, process: str, industry: str, erp: str) -> bytes:
    """Convert Markdown report text to a formatted PDF using reportlab."""
    from reportlab.lib.pagesizes import A4
    from reportlab.lib import colors
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import cm
    from reportlab.platypus import (
        SimpleDocTemplate, Paragraph, Spacer, HRFlowable,
        Table, TableStyle, PageBreak
    )
    from reportlab.lib.enums import TA_LEFT, TA_CENTER
    import io as _io

    buf = _io.BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=A4,
        leftMargin=2*cm, rightMargin=2*cm,
        topMargin=2.5*cm, bottomMargin=2*cm,
    )

    base = getSampleStyleSheet()

    sty = {
        "h1": ParagraphStyle("h1", parent=base["Heading1"],
                              fontSize=16, textColor=colors.HexColor("#1B3A6B"),
                              spaceAfter=10, spaceBefore=14),
        "h2": ParagraphStyle("h2", parent=base["Heading2"],
                              fontSize=13, textColor=colors.HexColor("#1B3A6B"),
                              spaceAfter=8, spaceBefore=12),
        "h3": ParagraphStyle("h3", parent=base["Heading3"],
                              fontSize=11, textColor=colors.HexColor("#2E5FA3"),
                              spaceAfter=6, spaceBefore=10),
        "normal": ParagraphStyle("normal", parent=base["Normal"],
                                 fontSize=9, leading=13, spaceAfter=4),
        "bullet": ParagraphStyle("bullet", parent=base["Normal"],
                                 fontSize=9, leading=13, leftIndent=14,
                                 bulletIndent=4, spaceAfter=3),
        "bold": ParagraphStyle("bold", parent=base["Normal"],
                               fontSize=9, leading=13, spaceAfter=4),
        "meta": ParagraphStyle("meta", parent=base["Normal"],
                               fontSize=9, textColor=colors.HexColor("#555555"),
                               spaceAfter=6),
    }

    story = []

    # Cover header
    story.append(Paragraph("RCM Design Effectiveness Review", sty["h1"]))
    story.append(Paragraph(
        f"<b>Process:</b> {process} &nbsp;&nbsp;|&nbsp;&nbsp; "
        f"<b>Industry:</b> {industry} &nbsp;&nbsp;|&nbsp;&nbsp; "
        f"<b>ERP:</b> {erp}",
        sty["meta"]
    ))
    story.append(HRFlowable(width="100%", thickness=1.5,
                             color=colors.HexColor("#1B3A6B"), spaceAfter=12))

    parsed = _parse_md_lines(md_text)

    # Group consecutive table_row items into tables
    i = 0
    while i < len(parsed):
        typ, content = parsed[i]

        if typ == "h1":
            story.append(Spacer(1, 6))
            story.append(Paragraph(content, sty["h1"]))
        elif typ == "h2":
            story.append(Spacer(1, 4))
            story.append(Paragraph(content, sty["h2"]))
        elif typ == "h3":
            story.append(Paragraph(content, sty["h3"]))
        elif typ == "hr":
            story.append(HRFlowable(width="100%", thickness=0.5,
                                     color=colors.HexColor("#AAAAAA"),
                                     spaceBefore=4, spaceAfter=4))
        elif typ == "bullet":
            story.append(Paragraph(f"• {content}", sty["bullet"]))
        elif typ == "bold_line":
            story.append(Paragraph(f"<b>{content}</b>", sty["bold"]))
        elif typ == "table_row":
            # Collect all consecutive table rows
            rows = [content]
            j = i + 1
            while j < len(parsed) and parsed[j][0] == "table_row":
                rows.append(parsed[j][1])
                j += 1
            i = j - 1

            if rows:
                # Build reportlab table
                col_count = max(len(r) for r in rows)
                # Pad short rows
                tdata = [r + [""] * (col_count - len(r)) for r in rows]

                col_width = (A4[0] - 4*cm) / col_count
                col_widths = [col_width] * col_count

                tbl = Table(tdata, colWidths=col_widths, repeatRows=1)
                tbl.setStyle(TableStyle([
                    ("BACKGROUND",   (0, 0), (-1, 0), colors.HexColor("#1B3A6B")),
                    ("TEXTCOLOR",    (0, 0), (-1, 0), colors.white),
                    ("FONTNAME",     (0, 0), (-1, 0), "Helvetica-Bold"),
                    ("FONTSIZE",     (0, 0), (-1, 0), 8),
                    ("FONTSIZE",     (0, 1), (-1, -1), 8),
                    ("FONTNAME",     (0, 1), (-1, -1), "Helvetica"),
                    ("ROWBACKGROUNDS", (0, 1), (-1, -1),
                     [colors.white, colors.HexColor("#EEF2F8")]),
                    ("GRID",         (0, 0), (-1, -1), 0.4, colors.HexColor("#AAAAAA")),
                    ("VALIGN",       (0, 0), (-1, -1), "TOP"),
                    ("TOPPADDING",   (0, 0), (-1, -1), 4),
                    ("BOTTOMPADDING",(0, 0), (-1, -1), 4),
                    ("LEFTPADDING",  (0, 0), (-1, -1), 4),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 4),
                    ("WORDWRAP",     (0, 0), (-1, -1), True),
                ]))
                story.append(Spacer(1, 4))
                story.append(tbl)
                story.append(Spacer(1, 6))
        elif typ == "normal":
            # Replace **bold** inline markers for reportlab
            text = content.replace("**", "<b>", 1)
            count = 0
            result = []
            for ch in content:
                if content[len(result):].startswith("**"):
                    tag = "<b>" if count % 2 == 0 else "</b>"
                    result.append(tag)
                    count += 1
            # Simpler approach: just strip ** for PDF normal text
            clean = content.replace("**", "")
            story.append(Paragraph(clean, sty["normal"]))

        i += 1

    doc.build(story)
    return buf.getvalue()


def generate_docx(md_text: str, process: str, industry: str, erp: str) -> bytes:
    """Convert Markdown report text to a formatted Word .docx using python-docx."""
    try:
        from docx import Document
        from docx.shared import Pt, RGBColor, Inches, Cm
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        from docx.oxml.ns import qn
        from docx.oxml import OxmlElement
        import io as _io
        import re as _re
    except ImportError:
        raise ImportError(
            "python-docx not installed. Run: pip install python-docx"
        )

    doc = Document()

    # ── Page margins
    for section in doc.sections:
        section.top_margin    = Cm(2.5)
        section.bottom_margin = Cm(2)
        section.left_margin   = Cm(2.5)
        section.right_margin  = Cm(2.5)

    # ── Cover header
    title_para = doc.add_heading("RCM Design Effectiveness Review", level=1)
    title_para.runs[0].font.color.rgb = RGBColor(0x1B, 0x3A, 0x6B)

    meta = doc.add_paragraph()
    meta.add_run(f"Process: ").bold = True
    meta.add_run(process)
    meta.add_run("   |   ")
    meta.add_run("Industry: ").bold = True
    meta.add_run(industry)
    meta.add_run("   |   ")
    meta.add_run("ERP: ").bold = True
    meta.add_run(erp)
    meta.paragraph_format.space_after = Pt(8)

    # horizontal rule via paragraph border
    def add_hr(doc):
        p = doc.add_paragraph()
        pPr = p._p.get_or_add_pPr()
        pBdr = OxmlElement("w:pBdr")
        bottom = OxmlElement("w:bottom")
        bottom.set(qn("w:val"), "single")
        bottom.set(qn("w:sz"), "12")
        bottom.set(qn("w:space"), "1")
        bottom.set(qn("w:color"), "1B3A6B")
        pBdr.append(bottom)
        pPr.append(pBdr)
        return p

    add_hr(doc)

    parsed = _parse_md_lines(md_text)

    def inline_bold(para, text):
        """Add a run to para, converting **text** to bold runs."""
        parts = _re.split(r"(\*\*.*?\*\*)", text)
        for part in parts:
            if part.startswith("**") and part.endswith("**"):
                run = para.add_run(part[2:-2])
                run.bold = True
            else:
                para.add_run(part)

    i = 0
    while i < len(parsed):
        typ, content = parsed[i]

        if typ == "h1":
            h = doc.add_heading(content, level=1)
            if h.runs:
                h.runs[0].font.color.rgb = RGBColor(0x1B, 0x3A, 0x6B)
        elif typ == "h2":
            h = doc.add_heading(content, level=2)
            if h.runs:
                h.runs[0].font.color.rgb = RGBColor(0x1B, 0x3A, 0x6B)
        elif typ == "h3":
            h = doc.add_heading(content, level=3)
            if h.runs:
                h.runs[0].font.color.rgb = RGBColor(0x2E, 0x5F, 0xA3)
        elif typ == "hr":
            add_hr(doc)
        elif typ == "bullet":
            p = doc.add_paragraph(style="List Bullet")
            inline_bold(p, content)
        elif typ == "bold_line":
            p = doc.add_paragraph()
            run = p.add_run(content)
            run.bold = True
        elif typ == "table_row":
            # Collect all consecutive table rows
            rows = [content]
            j = i + 1
            while j < len(parsed) and parsed[j][0] == "table_row":
                rows.append(parsed[j][1])
                j += 1
            i = j - 1

            if rows:
                col_count = max(len(r) for r in rows)
                tbl = doc.add_table(rows=len(rows), cols=col_count)
                tbl.style = "Table Grid"

                for r_idx, row_data in enumerate(rows):
                    row_cells = tbl.rows[r_idx].cells
                    for c_idx, cell_text in enumerate(row_data):
                        if c_idx < col_count:
                            cell = row_cells[c_idx]
                            cell.text = cell_text
                            # Header row styling
                            if r_idx == 0:
                                run = cell.paragraphs[0].runs
                                if run:
                                    run[0].bold = True
                                    run[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
                                # Blue background for header
                                tc_pr = cell._tc.get_or_add_tcPr()
                                shd = OxmlElement("w:shd")
                                shd.set(qn("w:val"), "clear")
                                shd.set(qn("w:color"), "auto")
                                shd.set(qn("w:fill"), "1B3A6B")
                                tc_pr.append(shd)
                doc.add_paragraph()  # spacing after table
        elif typ == "normal":
            p = doc.add_paragraph()
            inline_bold(p, content)

        i += 1

    buf = _io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


# ---------------------------------------------------------------------------
# Test API key
# ---------------------------------------------------------------------------
def test_api_key(api_key: str) -> bool:
    try:
        client = Groq(api_key=api_key)
        client.chat.completions.create(
            model="llama-3.1-8b-instant",
            messages=[{"role": "user", "content": "say hi"}],
            max_tokens=5,
        )
        return True
    except Exception:
        return False


# ---------------------------------------------------------------------------
# Main UI
# ---------------------------------------------------------------------------
if not api_key:
    st.info(
        "👈 Enter your **Groq API Key** in the sidebar to begin.\n\n"
        "Get a free key at 👉 https://console.groq.com/keys"
    )
    st.stop()

with st.spinner("🔑 Validating API key…"):
    if not test_api_key(api_key):
        st.error(
            "❌ Invalid or expired API key.\n\n"
            "Please go to https://console.groq.com/keys and create a new key."
        )
        st.stop()

st.success(f"✅ API key valid! Using model: `{selected_model}`")

# Show active context banner
st.info(
    f"📋 **Process:** {process_label}  |  "
    f"🏭 **Industry:** {industry_label}  |  "
    f"💻 **ERP:** {erp_label}  |  "
    f"🛠️ **Mode:** {tool_mode.split('  ')[1]}"
)

# ===========================================================================
# MODE A — Redraft & Classify Controls
# ===========================================================================
if "Redraft" in tool_mode:
    st.markdown("## ✏️ Redraft & Classify Controls")
    st.markdown(
        "Upload your RCM Excel file. Each control will be rewritten in **Big 4 audit style** "
        "and classified for **Key/Non-Key** and **Anti-Fraud** using ICFR/SOX criteria."
    )

    uploaded_file = st.file_uploader("📁 Upload RCM Excel File (.xlsx)", type=["xlsx"])

    if uploaded_file:
        df = pd.read_excel(uploaded_file)
        st.write("### 👀 Preview of Uploaded Data")
        st.dataframe(df.head(10))

        st.write("### 🔧 Column Configuration")
        col1, col2 = st.columns(2)
        with col1:
            control_col = st.selectbox(
                "Select the Control Activity Column:",
                df.columns,
                help="The column containing control descriptions to be redrafted"
            )
        with col2:
            risk_col = st.selectbox(
                "Select the Risk Column (optional):",
                ["None"] + list(df.columns),
                help="If selected, AI will use this risk description to improve its assessment"
            )

        st.info(
            "📌 **Control Type** is classified using ICFR/SOX Key Control criteria — "
            "management review controls over tax, financial statements, journal entries, "
            "provisions, reconciliations, bank payments, statutory returns, or significant "
            "estimates are always classified as **Key** unless a higher-level compensating "
            "control exists."
        )
        if risk_col != "None":
            st.success(
                f"✅ **Risk column selected:** `{risk_col}` — AI will use existing risk descriptions "
                "to improve Risk Rating, Anti-Fraud classification, and Key/Non-Key assessment."
            )

        if st.button("🚀 Process RCM", type="primary"):
            client        = Groq(api_key=api_key)
            all_results: dict[int, dict] = {}
            rows          = list(df.iterrows())
            batches       = [rows[i: i + batch_size] for i in range(0, len(rows), batch_size)]
            progress_bar  = st.progress(0)
            status_text   = st.empty()
            error_log     = []
            has_risk_col  = risk_col != "None"

            for b_idx, batch in enumerate(batches):
                status_text.text(f"Processing batch {b_idx + 1} of {len(batches)}…")

                controls_input = []
                for idx, row in batch:
                    item = {"original_id": int(idx), "description": str(row[control_col])}
                    if has_risk_col:
                        item["existing_risk"] = str(row[risk_col])
                    controls_input.append(item)

                try:
                    result = call_groq(
                        client, selected_model, controls_input,
                        max_retries, retry_delay,
                        has_risk_col=has_risk_col,
                        process=process_label,
                        industry=industry_label,
                        erp=erp_label,
                    )
                    for item in result.items:
                        control_type = apply_control_type_rule(
                            item.anti_fraud, item.risk_rating, item.control_type
                        )
                        all_results[item.original_id] = {
                            "Redrafted Control Activity": item.redrafted_control,
                            "Anti - Fraud (Y/N)":         item.anti_fraud,
                            "Risk Rating (L,M,H)":        item.risk_rating,
                            "Control Type (Key/Non-Key)": control_type,
                        }
                except Exception as e:
                    error_log.append(f"Batch {b_idx + 1}: {e}")
                    st.warning(f"⚠️ Batch {b_idx + 1} failed and was skipped.")

                progress_bar.progress((b_idx + 1) / len(batches))
                if b_idx < len(batches) - 1:
                    time.sleep(2)

            status_text.text("✅ Processing complete!")

            if error_log:
                with st.expander("⚠️ Errors (click to expand)"):
                    for err in error_log:
                        st.write(err)

            if all_results:
                res_df       = pd.DataFrame.from_dict(all_results, orient="index")
                cols_to_drop = [c for c in res_df.columns if c in df.columns]
                df_clean     = df.drop(columns=cols_to_drop)
                final_df     = pd.concat([df_clean, res_df], axis=1)

                st.success(f"🎉 Done! {len(all_results)} of {len(df)} rows processed.")

                key_count     = sum(1 for v in all_results.values() if v["Control Type (Key/Non-Key)"] == "Key")
                non_key_count = len(all_results) - key_count
                fraud_y_count = sum(1 for v in all_results.values() if v["Anti - Fraud (Y/N)"] == "Y")

                c1, c2, c3, c4 = st.columns(4)
                c1.metric("Total Processed",    len(all_results))
                c2.metric("Key Controls",        key_count)
                c3.metric("Non-Key Controls",    non_key_count)
                c4.metric("Anti-Fraud Controls", fraud_y_count)

                high = sum(1 for v in all_results.values() if v["Risk Rating (L,M,H)"].lower() == "high")
                med  = sum(1 for v in all_results.values() if v["Risk Rating (L,M,H)"].lower() == "medium")
                low  = sum(1 for v in all_results.values() if v["Risk Rating (L,M,H)"].lower() == "low")
                st.write("#### 📊 Risk Rating Breakdown")
                r1, r2, r3 = st.columns(3)
                r1.metric("🔴 High", high)
                r2.metric("🟡 Medium", med)
                r3.metric("🟢 Low", low)

                st.dataframe(final_df)

                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                    final_df.to_excel(writer, index=False, sheet_name="Redrafted RCM")
                st.download_button(
                    label="📥 Download Updated RCM Excel",
                    data=buffer.getvalue(),
                    file_name="Redrafted_RCM.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            else:
                st.error(
                    "❌ No rows were processed.\n"
                    "Try reducing batch size or increasing retry delay in the sidebar."
                )


# ===========================================================================
# MODE B — Full RCM Design Effectiveness Review
# ===========================================================================
else:
    st.markdown("## 🔍 Full RCM Design Effectiveness Review")
    st.markdown(
        "Upload your RCM Excel file. The AI will perform a comprehensive **14-part "
        "Design Effectiveness Review** benchmarked against COSO 2013, IFC (Companies Act 2013), "
        "ICAI Guidance Note, SOX Best Practices, and Anti-Fraud Framework."
    )

    if not selected_parts:
        st.warning("⚠️ Please select at least one report part in the sidebar.")

    uploaded_file = st.file_uploader("📁 Upload RCM Excel File (.xlsx)", type=["xlsx"])

    if uploaded_file:
        df = pd.read_excel(uploaded_file)
        st.write("### 👀 Preview of Uploaded RCM")
        st.dataframe(df.head(10))

        st.write("### 🔧 Select Columns to Include in Review")
        review_cols = st.multiselect(
            "Columns to pass to the AI for review:",
            list(df.columns),
            default=list(df.columns),
            help="Select all columns relevant to the review (risks, controls, owners, etc.)"
        )

        if not review_cols:
            st.warning("Please select at least one column.")
        else:
            parts_display = (
                "All 14 Parts"
                if "All Parts (Full Report)" in selected_parts
                else ", ".join(selected_parts)
            )
            st.info(
                f"📑 **Report scope:** {parts_display}\n\n"
                f"The AI will review **{len(df)} controls** across the selected parts. "
                f"Large RCMs may take 1–2 minutes."
            )

            if st.button("🔍 Run Design Effectiveness Review", type="primary"):
                client = Groq(api_key=api_key)

                # Convert the selected columns of the DataFrame to CSV text (no extra deps)
                review_df  = df[review_cols]
                rcm_text   = review_df.to_csv(index=True)

                review_prompt = build_design_review_prompt(
                    process=process_label,
                    industry=industry_label,
                    erp=erp_label,
                    parts=selected_parts,
                    rcm_text=rcm_text,
                )

                with st.spinner("🔍 AI is reviewing your RCM — this may take 1–2 minutes…"):
                    try:
                        review_output = call_groq_review(
                            client, selected_model, review_prompt,
                            int(max_retries), retry_delay,
                        )
                        st.success("✅ Design Effectiveness Review complete!")
                        st.markdown("---")
                        st.markdown("## 📋 RCM Design Effectiveness Review Report")
                        st.markdown(
                            f"**Process:** {process_label}  |  "
                            f"**Industry:** {industry_label}  |  "
                            f"**ERP:** {erp_label}"
                        )
                        st.markdown("---")
                        st.markdown(review_output)

                        # ── Download buttons ──────────────────────────
                        st.markdown("### 📥 Download Report")
                        dl1, dl2, dl3 = st.columns(3)

                        # 1) Markdown
                        with dl1:
                            st.download_button(
                                label="📄 Markdown (.md)",
                                data=review_output.encode("utf-8"),
                                file_name=f"RCM_Review_{process_label.replace(' ','_')}.md",
                                mime="text/markdown",
                                use_container_width=True,
                            )

                        # 2) Word (.docx)
                        with dl2:
                            try:
                                docx_bytes = generate_docx(
                                    review_output, process_label,
                                    industry_label, erp_label
                                )
                                st.download_button(
                                    label="📝 Word (.docx)",
                                    data=docx_bytes,
                                    file_name=f"RCM_Review_{process_label.replace(' ','_')}.docx",
                                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                    use_container_width=True,
                                )
                            except Exception as ex:
                                st.warning(f"Word export failed: {ex}")

                        # 3) PDF
                        with dl3:
                            try:
                                pdf_bytes = generate_pdf(
                                    review_output, process_label,
                                    industry_label, erp_label
                                )
                                st.download_button(
                                    label="📕 PDF (.pdf)",
                                    data=pdf_bytes,
                                    file_name=f"RCM_Review_{process_label.replace(' ','_')}.pdf",
                                    mime="application/pdf",
                                    use_container_width=True,
                                )
                            except Exception as ex:
                                st.warning(f"PDF export failed: {ex}")

                    except Exception as e:
                        st.error(f"❌ Review failed: {e}")
