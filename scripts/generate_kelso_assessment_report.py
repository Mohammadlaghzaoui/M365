# -*- coding: utf-8 -*-
"""
Builds the consolidated, enterprise-level Kelso Industries M365 Security &
Tenant-to-Tenant Readiness Assessment Word report from the source assessments.
"""
import os
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.section import WD_SECTION
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

OUT = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                   "Reports", "Kelso_M365_Security_Assessment_Enterprise_v1.0.docx")

# Corporate palette
NAVY = "1F3864"
BLUE = "2F5597"
LIGHT = "D6E4F0"
GREY = "F2F2F2"
RED = "C00000"
AMBER = "ED7D31"
GOLD = "FFC000"
GREEN = "538135"

RAG_FILL = {"Red": "F4CCCC", "Amber": "FCE5CD", "Green": "D9EAD3",
            "Critical": "F4CCCC", "High": "FCE5CD", "Medium-High": "FFF2CC",
            "Medium": "FFF2CC", "Low": "D9EAD3"}
RAG_TEXT = {"Red": RED, "Amber": "B45F06", "Green": GREEN,
            "Critical": RED, "High": "B45F06", "Medium-High": "B45F06",
            "Medium": "7F6000", "Low": GREEN}


def shade(cell, color):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:fill'), color)
    tcPr.append(shd)


def cell_text(cell, text, bold=False, color=None, size=9.5, align=None):
    cell.text = ""
    p = cell.paragraphs[0]
    run = p.add_run(text)
    run.font.size = Pt(size)
    run.font.name = "Calibri"
    run.bold = bold
    if color:
        run.font.color.rgb = RGBColor.from_string(color)
    if align:
        p.alignment = align
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(2)


def make_table(doc, rows, cols, widths=None, header_fill=BLUE):
    t = doc.add_table(rows=rows, cols=cols)
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = False
    if widths:
        for r in t.rows:
            for i, w in enumerate(widths):
                r.cells[i].width = Cm(w)
    return t


def header_cells(table, labels, fill=BLUE):
    for i, lab in enumerate(labels):
        cell = table.rows[0].cells[i]
        cell_text(cell, lab, bold=True, color="FFFFFF", size=10)
        shade(cell, fill)


def fill_row(table, ridx, values, rag_cols=()):
    for i, v in enumerate(values):
        cell = table.rows[ridx].cells[i]
        if i in rag_cols and v in RAG_FILL:
            cell_text(cell, v, bold=True, color=RAG_TEXT[v],
                      align=WD_ALIGN_PARAGRAPH.CENTER)
            shade(cell, RAG_FILL[v])
        else:
            cell_text(cell, v)
        if ridx % 2 == 0 and i not in rag_cols:
            shade(cell, GREY)


def add_toc(doc):
    p = doc.add_paragraph()
    run = p.add_run()
    fld1 = OxmlElement('w:fldChar'); fld1.set(qn('w:fldCharType'), 'begin')
    instr = OxmlElement('w:instrText'); instr.set(qn('xml:space'), 'preserve')
    instr.text = 'TOC \\o "1-2" \\h \\z \\u'
    fld2 = OxmlElement('w:fldChar'); fld2.set(qn('w:fldCharType'), 'separate')
    t = OxmlElement('w:t')
    t.text = "Table of contents — right-click and choose 'Update Field' in Word to populate."
    fld3 = OxmlElement('w:fldChar'); fld3.set(qn('w:fldCharType'), 'end')
    run._r.append(fld1); run._r.append(instr); run._r.append(fld2)
    run._r.append(t); run._r.append(fld3)


def add_page_number_footer(section):
    footer = section.footer
    p = footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run("Kelso Industries — Microsoft 365 Security & Tenant-to-Tenant Readiness Assessment  |  Confidential  |  Page ")
    run.font.size = Pt(8); run.font.color.rgb = RGBColor.from_string("808080")
    run2 = p.add_run()
    fld1 = OxmlElement('w:fldChar'); fld1.set(qn('w:fldCharType'), 'begin')
    instr = OxmlElement('w:instrText'); instr.set(qn('xml:space'), 'preserve')
    instr.text = 'PAGE'
    fld2 = OxmlElement('w:fldChar'); fld2.set(qn('w:fldCharType'), 'end')
    run2._r.append(fld1); run2._r.append(instr); run2._r.append(fld2)
    run2.font.size = Pt(8); run2.font.color.rgb = RGBColor.from_string("808080")


def style_doc(doc):
    st = doc.styles["Normal"]
    st.font.name = "Calibri"; st.font.size = Pt(10.5)
    st.paragraph_format.space_after = Pt(6)
    for name, size, color, before in (
            ("Heading 1", 16, NAVY, 18), ("Heading 2", 13, BLUE, 12),
            ("Heading 3", 11.5, BLUE, 8)):
        h = doc.styles[name]
        h.font.name = "Calibri"; h.font.size = Pt(size); h.font.bold = True
        h.font.color.rgb = RGBColor.from_string(color)
        h.paragraph_format.space_before = Pt(before)
        h.paragraph_format.space_after = Pt(6)
        h.font.element.rPr.rFonts.set(qn('w:eastAsia'), 'Calibri')
    title = doc.styles["Title"]
    title.font.name = "Calibri"; title.font.size = Pt(28); title.font.bold = True
    title.font.color.rgb = RGBColor.from_string(NAVY)


def bullets(doc, items):
    for it in items:
        p = doc.add_paragraph(it, style="List Bullet")
        p.paragraph_format.space_after = Pt(2)
        for r in p.runs:
            r.font.size = Pt(10.5)


def para(doc, text, bold=False, size=10.5, color=None, align=None, italic=False):
    p = doc.add_paragraph()
    r = p.add_run(text)
    r.bold = bold; r.italic = italic
    r.font.size = Pt(size)
    if color:
        r.font.color.rgb = RGBColor.from_string(color)
    if align:
        p.alignment = align
    return p


# =====================================================================
doc = Document()
style_doc(doc)
sec = doc.sections[0]
sec.top_margin = Cm(2.2); sec.bottom_margin = Cm(2.2)
sec.left_margin = Cm(2.2); sec.right_margin = Cm(2.2)
add_page_number_footer(sec)

# ---------------- COVER PAGE ----------------
for _ in range(4):
    doc.add_paragraph()
para(doc, "MICROSOFT 365", bold=True, size=14, color=BLUE,
     align=WD_ALIGN_PARAGRAPH.CENTER)
p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run("Security & Tenant-to-Tenant\nReadiness Assessment")
r.bold = True; r.font.size = Pt(30); r.font.color.rgb = RGBColor.from_string(NAVY)
para(doc, "Kelso Industries", bold=True, size=18, color=AMBER,
     align=WD_ALIGN_PARAGRAPH.CENTER)
doc.add_paragraph()
para(doc, "Consolidated Enterprise Report — Final Client Version", size=12,
     color="595959", align=WD_ALIGN_PARAGRAPH.CENTER)
for _ in range(6):
    doc.add_paragraph()

t = make_table(doc, 6, 2, widths=[5.5, 10.5])
meta = [("Prepared for", "Kelso Industries"),
        ("Prepared by", "Mohamed Laghzaoui — Senior M365 Engineer"),
        ("Assessment date", "12 June 2026"),
        ("Document version", "1.0 (Consolidated — supersedes all prior drafts)"),
        ("Classification", "Confidential — authorized Kelso Industries stakeholders only"),
        ("Evidence basis", "Microsoft 365 / Entra / Purview / Intune / Defender portal observations, consultant-reviewed and revalidated")]
for i, (k, v) in enumerate(meta):
    cell_text(t.rows[i].cells[0], k, bold=True, color="FFFFFF", size=10)
    shade(t.rows[i].cells[0], NAVY)
    cell_text(t.rows[i].cells[1], v, size=10)
doc.add_paragraph()
para(doc, "Confidentiality notice: this document contains administrative security observations "
          "about the Kelso Industries Microsoft 365 tenant and must be shared only with "
          "authorized stakeholders.", size=8.5, color="808080", italic=True,
     align=WD_ALIGN_PARAGRAPH.CENTER)
doc.add_page_break()

# ---------------- DOCUMENT CONTROL ----------------
doc.add_heading("Document Control", level=1)
t = make_table(doc, 3, 4, widths=[2.2, 3.2, 5.6, 5.0])
header_cells(t, ["Version", "Date", "Description", "Author"])
fill_row(t, 1, ["0.x", "May–June 2026",
                "Working drafts: portal assessment, reference-practice review, consultant review, live portal revalidation",
                "Mohamed Laghzaoui"])
fill_row(t, 2, ["1.0", "12 June 2026",
                "Consolidated enterprise report — all findings merged, validated wording, client-ready",
                "Mohamed Laghzaoui"])
doc.add_paragraph()
doc.add_heading("Table of Contents", level=1)
add_toc(doc)
doc.add_page_break()

# ---------------- 1. EXECUTIVE SUMMARY ----------------
doc.add_heading("1. Executive Summary", level=1)
para(doc, "This assessment reviews the Microsoft 365 security and governance posture of Kelso "
          "Industries from a senior M365 engineering perspective, and evaluates the tenant's "
          "readiness to act as the target environment for upcoming tenant-to-tenant migrations. "
          "The tenant shows meaningful adoption of Microsoft 365 services, including Exchange "
          "Online, Microsoft Teams, SharePoint, OneDrive, Microsoft Entra, Intune, Purview and "
          "Microsoft 365 Copilot.")
para(doc, "The most important conclusion is that the Microsoft 365 environment is not "
          "fundamentally weak; it is in a transitional maturity stage. Several controls exist or "
          "are partially available, but they need to be completed, governed, documented and "
          "measured. The key theme is not missing capability — it is that important controls are "
          "not yet fully operationalized, validated or governed end-to-end.", bold=False)

t = make_table(doc, 2, 3, widths=[4.0, 6.0, 6.0])
header_cells(t, ["Current Risk Level", "Priority Findings", "Overall Direction"])
fill_row(t, 1, ["Medium",
                "Critical findings exist in privileged access and authentication method governance.",
                "Improve control maturity without disrupting business adoption."], rag_cols=())
shade(t.rows[1].cells[0], RAG_FILL["Medium"])

doc.add_heading("1.1 Positive Observations", level=2)
bullets(doc, [
    "Microsoft 365 is actively used across the organization: 742 active users in the admin dashboard and 459 active Microsoft 365 users across apps.",
    "Microsoft Entra ID P1 is present, providing a base for Conditional Access and identity governance improvements.",
    "Conditional Access is already in place: 4 Microsoft-managed policies and 4 user-created policies.",
    "Purview Audit search is available and supports audit investigation across Microsoft 365 workloads.",
    "Intune is in use for Windows device management, with 279 Windows devices enrolled.",
    "SharePoint Advanced Management is active, enabling oversharing and site-lifecycle assessments.",
    "Microsoft 365 Copilot is actively deployed: 42 assigned licenses and visible adoption (~5.3K prompts).",
    "Exchange admin center is surfacing actionable mail-flow insights, including forwarding activity that can be investigated.",
])

doc.add_heading("1.2 Executive Domain Dashboard", level=2)
t = make_table(doc, 7, 4, widths=[3.2, 1.8, 2.0, 9.0])
header_cells(t, ["Domain", "Status", "Priority", "Consultant assessment"])
rows = [
    ("Identity", "Red", "Critical", "Privileged identity governance and authentication method enforcement require immediate attention. PIM is not operationalized (Entra ID P1)."),
    ("Collaboration", "Amber", "High", "Teams, SharePoint and OneDrive are operational, but external sharing defaults and external access are broad and require governance."),
    ("Endpoint", "Amber", "High", "Intune is in use for Windows devices; policy coverage and the 'Grant Local Admin' policy require detailed validation."),
    ("Data Protection", "Amber", "High", "Purview DLP is configured in test/simulation mode and not yet operating as an enforced control."),
    ("Governance", "Amber", "Medium", "Service governance requires improvement, including ownership of Message Center (744 unread), agents and enterprise applications."),
    ("Overall Maturity", "Amber", "Medium", "The platform foundation is present, but operating controls must be completed, validated and evidenced."),
]
for i, r in enumerate(rows, 1):
    fill_row(t, i, list(r), rag_cols=(1, 2))

doc.add_page_break()

# ---------------- 2. TENANT SNAPSHOT ----------------
doc.add_heading("2. Tenant & Platform Snapshot", level=1)
para(doc, "The tenant discovery indicates an established Microsoft 365 environment with broad "
          "adoption and a meaningful population of users, groups, applications, devices and "
          "collaboration workloads.")
t = make_table(doc, 14, 2, widths=[7.0, 9.0])
header_cells(t, ["Metric", "Observed value"])
snap = [
    ("Primary domain", "kelso-industries.com"),
    ("Tenant ID", "Masked in this client-facing version"),
    ("Entra users / groups", "1,279 users / 305 groups"),
    ("Active users (M365 admin dashboard)", "742"),
    ("Active Microsoft 365 users across apps", "459"),
    ("Workload usage (active users)", "Exchange 271 · Teams 213 · SharePoint 116 · OneDrive 112"),
    ("Microsoft 365 licenses", "696 of 706 assigned"),
    ("Microsoft 365 Copilot licenses", "42 of 42 assigned (~5.3K prompts used)"),
    ("Agents", "247 total agents observed"),
    ("Entra licensing", "Microsoft Entra ID P1 (PIM requires P2 / Entra ID Governance)"),
    ("Conditional Access policies", "4 Microsoft-managed + 4 user-created"),
    ("Intune devices", "279 Windows devices; no Apple/Android/Linux population visible"),
    ("Enterprise applications", "152 observed (including MigrationWiz)"),
]
for i, (k, v) in enumerate(snap, 1):
    fill_row(t, i, [k, v])
para(doc, "Recommendations: maintain a formal tenant inventory covering identities, groups, "
          "applications, devices, workloads and administrative roles; define service ownership "
          "for Microsoft 365 workloads, enterprise applications, agents and operational health "
          "reviews.", italic=True, size=10)

# ---------------- 3. LICENSING ----------------
doc.add_heading("3. Licensing Assessment", level=1)
para(doc, "Licensing is sufficient for several Microsoft 365 capabilities, but some security and "
          "governance controls require licensing validation before implementation decisions are made.")
bullets(doc, [
    "Conduct a cost-benefit assessment of Microsoft Entra ID P2 or Entra ID Governance for privileged identities — this is the strongest single investment case identified.",
    "Confirm whether Microsoft Defender and Purview licensing is sufficient for the desired security, compliance and reporting scope.",
    "Align Copilot licensing with data governance readiness, especially SharePoint and OneDrive permission hygiene.",
])

# ---------------- 4. DETAILED FINDINGS ----------------
doc.add_page_break()
doc.add_heading("4. Detailed Findings & Recommended Improvements", level=1)
para(doc, "Each finding below lists severity, control owner, the evidence observed during the "
          "assessment, the business risk and the recommended remediation actions. Where portal "
          "state could not be fully confirmed, the finding is explicitly marked as requiring "
          "validation rather than assumed.")

FINDINGS = [
    ("F1. Privileged Identity Management is not operationalized", "Critical",
     "Identity / Security", "Immediate executive priority",
     "The PIM portal indicates that Microsoft Entra ID P2 or Entra ID Governance is required; the tenant currently runs Entra ID P1, so PIM is not operationalized.",
     "Administrators may hold standing privilege rather than just-in-time, approved and time-limited access. This increases the impact of admin compromise.",
     ["Confirm licensing for Entra ID P2 or Entra ID Governance for all privileged users.",
      "Enable PIM for Microsoft Entra roles.",
      "Convert permanent privileged role assignments to eligible assignments where possible.",
      "Require MFA, justification, approval and limited activation duration for high-impact roles.",
      "Document break-glass accounts and monitor their usage."]),
    ("F2. Privileged role assignments exceed recommended threshold", "High",
     "Identity / IT Leadership", "High priority",
     "The Entra roles page reports 18 privileged role assignments; portal guidance recommends limiting privileged assignments.",
     "Excess privileged assignments increase attack surface and reduce accountability. Privileged roles should be limited, justified and reviewed.",
     ["Export all role assignments and identify permanent privileged assignments.",
      "Remove stale, duplicate or unnecessary privileged roles.",
      "Replace Global Administrator assignments with least-privilege workload roles where possible.",
      "Create a privileged access register with owner, reason and review status.",
      "Perform quarterly privileged access reviews."]),
    ("F3. Authentication method governance is incomplete", "High",
     "Identity / Security", "High priority",
     "Authentication Methods migration is in progress. Microsoft Authenticator adoption and enforcement require validation. Temporary Access Pass is enabled for all users; Email OTP is enabled.",
     "MFA may be inconsistent or dependent on legacy settings. TAP enabled for all users should be tightened to controlled scenarios.",
     ["Complete migration to the unified Authentication Methods policy.",
      "Enable Microsoft Authenticator with number matching for users or phased security groups.",
      "Restrict Temporary Access Pass to approved onboarding/helpdesk groups, with short lifetime and one-time use.",
      "Introduce phishing-resistant MFA (FIDO2/passkeys) for administrators and sensitive roles.",
      "Export MFA registration status for users and admins."]),
    ("F4. Conditional Access exists but requires validation", "High",
     "Identity / Security", "High priority",
     "Conditional Access shows 4 Microsoft-managed and 4 user-created policies. Detailed policy design, exclusions and coverage were not validated.",
     "The presence of policies does not prove complete coverage. Risk remains if exclusions are broad, policies are report-only, legacy authentication is not blocked or admins are not fully covered.",
     ["Export all Conditional Access policies with assignments, exclusions, apps, conditions, grant and session controls.",
      "Confirm admin MFA covers all privileged roles.",
      "Confirm legacy authentication is blocked.",
      "Confirm break-glass accounts are excluded only where required and monitored.",
      "Run What-If testing for admins, users, guests, unmanaged devices and external networks."]),
    ("F5. Data Loss Prevention is not yet mature", "High",
     "Compliance / Security", "High priority",
     "Purview shows one visible DLP policy (the default Copilot DLP policy) configured in test/simulation mode and not enforced.",
     "A single test-mode DLP policy does not provide mature protection for Exchange, SharePoint, OneDrive, Teams and endpoint scenarios. Copilot adoption increases the importance of DLP enforcement.",
     ["Define sensitive information categories relevant to the business.",
      "Create DLP policies for Exchange, SharePoint, OneDrive and Teams.",
      "Tune in test mode, then move mature policies to enforcement via a controlled change process.",
      "Enable policy tips, user notifications and incident reporting.",
      "Align DLP with sensitivity labels and Copilot governance."]),
    ("F6. Sensitivity label posture requires confirmation", "Medium-High",
     "Compliance / Information Governance", "High priority",
     "The sensitivity labels configuration could not be confirmed as mature from the portal during this assessment.",
     "Without verified sensitivity labels, the organization lacks a consistent classification foundation for DLP, Teams/site controls and Copilot readiness.",
     ["Export labels and label policies from Purview.",
      "Confirm a clear label model such as Public, Internal, Confidential and Highly Confidential.",
      "Confirm labels are published to the right users and groups.",
      "Enable container labels for Teams and SharePoint sites where external sharing must be restricted.",
      "Use labels to support DLP and Copilot governance."]),
    ("F7. Enterprise application and SSO governance require formal review", "High",
     "Identity / Application Owners", "High priority",
     "The tenant contains 152 enterprise applications, including business-critical integrations and MigrationWiz.",
     "Large app estates can accumulate stale service principals, excessive permissions, ownerless applications, unmanaged consent and expired credentials — common identity and data-access risks. The number of applications is not itself a problem; the risk is the absence of an operating model.",
     ["Export all enterprise applications with owners, sign-in activity, permissions, credentials and certificate expiry.",
      "Remove stale or ownerless applications after validation.",
      "Review high-risk permissions such as Mail.Read, Files.Read.All, Sites.Read.All and Directory.ReadWrite.All.",
      "Enable or validate the admin consent workflow; limit user consent.",
      "Assign business and technical owners to every critical application; prefer group-based SSO assignments."]),
    ("F8. SharePoint and OneDrive sharing defaults are too permissive", "High",
     "SharePoint / Information Governance", "High priority",
     "Default file/folder sharing uses 'Anyone with the link' with Edit permission and no observed expiration.",
     "Anonymous edit links can be forwarded, persist longer than intended and expose content without identity assurance. This is a major risk in a Copilot-enabled tenant.",
     ["Set default link type to 'Specific people'.",
      "Set default permission to View.",
      "Require expiration for Anyone links or disable Anyone links by default.",
      "Restrict anonymous links to approved sites only.",
      "Run a SharePoint Advanced Management oversharing assessment."]),
    ("F9. Teams external collaboration is too broad", "High",
     "Teams / Security", "High priority",
     "Teams external access allows all external domains; guest access is enabled and multiple third-party storage providers are enabled.",
     "Broad external collaboration increases phishing, social engineering and data leakage exposure, and makes governance harder when combined with permissive file sharing.",
     ["Review whether 'Allow all external domains' is required; consider an approved-domain model for high-risk scenarios.",
      "Disable unmanaged Microsoft account and trial tenant communication unless business-approved.",
      "Review and disable third-party storage integrations that are not approved.",
      "Implement guest access reviews and owner attestation.",
      "Use sensitivity labels to restrict external access for confidential teams and sites."]),
    ("F10. Exchange forwarding activity requires investigation", "High",
     "Exchange / Security Operations", "Immediate operational priority",
     "Exchange reporting shows 7 auto-forwarded messages and 1 new remote domain receiving forwarded email.",
     "External forwarding may be legitimate, but it is also a common mailbox compromise and data exfiltration method.",
     ["Identify source mailboxes and destination domain for the observed forwarding activity.",
      "Validate business justification with mailbox owners.",
      "Block automatic external forwarding by default through the outbound spam policy.",
      "Review inbox rules, transport rules and mailbox forwarding settings.",
      "Create alerts for new forwarding rules and new forwarding domains."]),
    ("F11. Endpoint management coverage needs validation", "Medium-High",
     "Endpoint / Intune", "Medium-high priority",
     "Intune shows 279 Windows devices, 2 compliance policies and 2 configuration policies, including a 'Grant Local Admin' policy. No non-Windows device population is visible.",
     "Endpoint management appears Windows-focused. If mobile or BYOD access exists, data may be accessible without sufficient device or app protection. The Grant Local Admin policy requires scope review.",
     ["Review the Grant Local Admin policy scope, target groups and business justification.",
      "Validate Windows Security Baseline settings (BitLocker, Defender, firewall, updates, compliance).",
      "Confirm whether mobile/BYOD access is allowed; if so, implement Intune App Protection Policies.",
      "Use device compliance state in Conditional Access where appropriate."]),
    ("F12. Copilot and agents require governance alignment", "High",
     "M365 Governance / AI Governance", "High priority",
     "The dashboard shows 42 assigned Copilot licenses, ~5.3K prompts used and 247 total agents.",
     "Copilot and agents do not create permissions, but they make over-permissioned content easier to discover. They increase the importance of permission hygiene, DLP, labels and app governance.",
     ["Run oversharing and inactive-site assessments before expanding Copilot.",
      "Remediate sites with anonymous links, organization-wide access, broken inheritance and missing owners.",
      "Create an agent inventory with owner, purpose, data access scope and approval status.",
      "Align Copilot rollout with DLP, sensitivity labels and access reviews.",
      "Report Copilot readiness to IT leadership."]),
]

for title, sev, owner, prio, evidence, risk, actions in FINDINGS:
    doc.add_heading(title, level=2)
    t = make_table(doc, 2, 3, widths=[4.2, 6.0, 5.8])
    header_cells(t, ["Severity", "Control owner", "Improvement priority"], fill=NAVY)
    fill_row(t, 1, [sev, owner, prio], rag_cols=(0,))
    para(doc, "Observed evidence: " + evidence, size=10)
    para(doc, "Risk and business impact: " + risk, size=10)
    para(doc, "Recommended actions:", bold=True, size=10)
    bullets(doc, actions)

# ---------------- 5. WORKLOAD REVIEWS ----------------
doc.add_page_break()
doc.add_heading("5. Workload Assessments", level=1)

WORKLOADS = [
    ("5.1 Exchange Online", [
        ("Exchange usage", "271 active Exchange users observed in workload reporting."),
        ("Auto-forwarding", "7 auto-forwarded messages observed in reporting."),
        ("Remote forwarding domains", "1 new remote domain receiving forwarded email observed."),
        ("Transport rules / mailbox audit", "Requires validation."),
    ], ["Review mailbox forwarding, inbox rules and transport rules.",
        "Block automatic external forwarding unless explicitly approved and documented.",
        "Create alerting for new forwarding rules and unusual mailbox access patterns."]),
    ("5.2 Email Authentication (SPF / DKIM / DMARC)", [
        ("SPF / DKIM / DMARC", "Not fully validated during this assessment — no conclusion should be made until DNS records and Exchange Online configuration are reviewed."),
        ("Accepted domains", "Requires validation."),
    ], ["Validate SPF includes only approved sending services.",
        "Confirm DKIM is enabled for all production mail domains.",
        "Review DMARC policy and reporting configuration.",
        "Align third-party senders with approved DNS and mail authentication standards."]),
    ("5.3 Microsoft Teams", [
        ("Teams usage", "213 active Teams users observed."),
        ("External access", "Configured to allow all external domains."),
        ("Guest access", "Enabled; guest interaction and external collaboration features are broad."),
        ("Third-party storage", "Multiple third-party storage providers enabled."),
    ], ["Move external access toward a managed allow/block model where appropriate.",
        "Validate guest access settings against business requirements.",
        "Disable third-party storage services that are not approved.",
        "Implement periodic review of external users and Teams with guests."]),
    ("5.4 SharePoint & OneDrive", [
        ("Usage", "116 active SharePoint users and 112 active OneDrive users observed."),
        ("Default sharing link", "'Anyone with the link' with Edit permission observed."),
        ("Anonymous link expiration", "Not confirmed during this assessment."),
    ], ["Change default sharing links to 'Specific people' or 'People in your organization'.",
        "Set default permission to View unless business need requires Edit.",
        "Validate anonymous link expiration and site-level external sharing exceptions.",
        "Review permissions before expanding Copilot usage — Copilot inherits existing Microsoft 365 permissions."]),
    ("5.5 Intune & Endpoint", [
        ("Windows devices", "279 observed; no Apple, macOS, Android or Linux population visible."),
        ("Compliance / configuration policies", "2 compliance and 2 configuration policies, including Windows Security Baseline and 'Grant Local Admin'."),
    ], ["Review compliance policies and assignments against Windows security baseline requirements.",
        "Validate Defender, BitLocker, firewall, update, local administrator and compliance posture.",
        "Review the Grant Local Admin policy, target groups and business justification.",
        "Use device compliance state in Conditional Access where appropriate."]),
    ("5.6 Microsoft Defender", [
        ("Portal state", "Defender portal access was validated; it displayed a Defender for Business onboarding/setup experience rather than a completed Secure Score review."),
        ("Defender for Office 365 / Endpoint / Cloud Apps", "Not confirmed during this assessment."),
    ], ["Validate enabled Defender workloads and licensing.",
        "Review anti-phishing, Safe Links, Safe Attachments and impersonation protection if Defender for Office 365 is licensed.",
        "Validate endpoint onboarding and EDR health if Defender for Endpoint is licensed.",
        "Use Microsoft Secure Score as an input, not as the only security maturity metric."]),
    ("5.7 Purview, Compliance & Service Governance", [
        ("DLP", "Default Copilot DLP policy observed in test/simulation mode; 'Policy not enforced' displayed."),
        ("Sensitivity labels", "Not confirmed during this assessment."),
        ("Audit", "Audit search capability observed."),
        ("Message Center", "744 unread Message Center messages observed — service governance requires ownership."),
    ], ["Define DLP policies for Exchange, SharePoint, OneDrive, Teams, endpoint and Copilot scenarios where required.",
        "Move validated DLP scenarios from simulation to enforcement using a controlled change process.",
        "Validate sensitivity labels, label publication policies and encryption requirements.",
        "Assign ownership for Message Center, service health and roadmap impact reviews."]),
]
for title, state_rows, recs in WORKLOADS:
    doc.add_heading(title, level=2)
    t = make_table(doc, len(state_rows) + 1, 2, widths=[6.0, 10.0])
    header_cells(t, ["Area", "Observed assessment"])
    for i, kv in enumerate(state_rows, 1):
        fill_row(t, i, list(kv))
    para(doc, "Recommendations:", bold=True, size=10)
    bullets(doc, recs)

# ---------------- 6. T2T READINESS ----------------
doc.add_page_break()
doc.add_heading("6. Hybrid & Tenant-to-Tenant Migration Readiness", level=1)
para(doc, "Kelso Industries is the target tenant for a multi-company tenant-to-tenant migration "
          "program. The readiness items below must be validated and closed before migration "
          "execution; they directly affect identity strategy, mail routing and cutover planning.")
T2T = [
    ("6.1 Hybrid Identity & Entra Connect",
     "Hybrid identity, Entra Connect / Cloud Sync status and synchronization health were not confirmed during this assessment.",
     ["Confirm whether Kelso uses cloud-only identity or hybrid identity.",
      "Validate synchronization tool, health status and sync error backlog.",
      "Confirm source anchor, UPN routing, proxyAddresses and mail attributes.",
      "Resolve sync and attribute issues before migration execution."]),
    ("6.2 Exchange Hybrid",
     "Exchange hybrid configuration was not confirmed during the assessment.",
     ["Validate whether Exchange hybrid is present and still required.",
      "Review connectors, accepted domains, remote domains and mail routing.",
      "Confirm coexistence requirements before any tenant-to-tenant migration."]),
    ("6.3 MigrationWiz Readiness",
     "MigrationWiz was observed as an enterprise application, but migration readiness was not fully validated.",
     ["Validate MigrationWiz application permissions and consent model.",
      "Confirm migration scope: workloads, identities, domains, mailboxes, OneDrive, SharePoint and Teams.",
      "Document prerequisites, throttling approach, coexistence requirements and rollback plan."]),
    ("6.4 Tenant-to-Tenant Readiness",
     "Tenant-to-tenant readiness requires additional discovery before a migration plan can be finalized.",
     ["Create a full migration inventory covering users, groups, mailboxes, aliases, domains, devices, OneDrive, SharePoint, Teams and enterprise apps for each source tenant.",
      "Validate source and target identity strategy, UPN/domain plan and coexistence requirements.",
      "Prepare a tenant-to-tenant cutover and rollback approach per migration wave.",
      "Close the Critical identity findings (F1–F3) before onboarding migrated users, so new accounts land in a governed environment."]),
]
for title, state, recs in T2T:
    doc.add_heading(title, level=2)
    para(doc, state, size=10)
    para(doc, "Recommendations:", bold=True, size=10)
    bullets(doc, recs)

# ---------------- 7. RISK REGISTER ----------------
doc.add_page_break()
doc.add_heading("7. Risk Register", level=1)
t = make_table(doc, 11, 5, widths=[1.2, 5.6, 1.6, 1.8, 5.8])
header_cells(t, ["ID", "Risk", "Status", "Priority", "Recommended action"])
RISKS = [
    ("R1", "Privileged identity governance is not fully operationalized.", "Red", "Critical",
     "Assess Entra ID P2/Governance and implement PIM for privileged roles."),
    ("R2", "Authentication method enforcement requires validation.", "Red", "Critical",
     "Complete methods migration and validate Authenticator adoption and enforcement."),
    ("R3", "Conditional Access exists but requires detailed validation.", "Amber", "High",
     "Export, review and test policies before relying on coverage."),
    ("R4", "DLP is in test/simulation mode.", "Amber", "High",
     "Move validated DLP policies toward enforcement."),
    ("R5", "SharePoint/OneDrive sharing defaults are permissive.", "Amber", "High",
     "Review default link type, permissions and expiration."),
    ("R6", "Teams external collaboration is broad.", "Amber", "High",
     "Review guest and external collaboration controls."),
    ("R7", "Exchange forwarding activity requires review.", "Amber", "High",
     "Investigate forwarding and block unapproved external forwarding."),
    ("R8", "Enterprise app governance requires lifecycle control.", "Amber", "Medium",
     "Inventory app owners, permissions, consent and usage."),
    ("R9", "Service governance requires improvement.", "Amber", "Medium",
     "Assign owners for Message Center, service health and roadmap reviews."),
    ("R10", "Tenant-to-tenant readiness requires further validation.", "Amber", "Medium",
     "Complete migration discovery and dependency mapping."),
]
for i, r in enumerate(RISKS, 1):
    fill_row(t, i, list(r), rag_cols=(2, 3))

# ---------------- 8. TARGET STATE ----------------
doc.add_heading("8. Recommended Control Target State", level=1)
t = make_table(doc, 13, 2, widths=[5.0, 11.0])
header_cells(t, ["Control area", "Target state"])
TARGETS = [
    ("Privileged access", "PIM enabled and used for just-in-time privileged role activation with approval, MFA and justification."),
    ("Admin roles", "Privileged assignments reduced, justified and reviewed regularly via a privileged access register."),
    ("Authentication methods", "Unified Authentication Methods migration complete; strong MFA enforced; phishing-resistant methods for admins."),
    ("Conditional Access", "Documented CA matrix covering admins, users, guests, legacy auth, unmanaged devices and sensitive apps."),
    ("SSO / Enterprise apps", "All applications have owners, reviewed permissions, sign-in review, consent governance and credential monitoring."),
    ("DLP", "DLP policies enforced across Exchange, SharePoint, OneDrive and Teams with incident handling."),
    ("Sensitivity labels", "Labels published and used for information protection, Teams/site containers and Copilot governance."),
    ("SharePoint / OneDrive", "Default links are 'Specific people' with View permission; anonymous links restricted and expiring."),
    ("Teams collaboration", "External access and guests governed by policy, labels and owner attestation with periodic review."),
    ("Exchange", "External forwarding blocked by default with approved, documented and monitored exceptions."),
    ("Endpoint", "Windows baseline enforced; local admin controlled; mobile/BYOD strategy defined with app protection."),
    ("AI / Copilot", "Copilot and agent estate inventoried, owned, permission-hygienic and aligned with DLP and labels."),
]
for i, kv in enumerate(TARGETS, 1):
    fill_row(t, i, list(kv))

# ---------------- 9. ROADMAP ----------------
doc.add_page_break()
doc.add_heading("9. Remediation Roadmap", level=1)
para(doc, "The roadmap is organized by priority and operational dependency rather than fixed "
          "dates, keeping it appropriate for audit presentation while giving management a "
          "practical path forward.")
t = make_table(doc, 7, 3, widths=[4.4, 4.6, 7.0])
header_cells(t, ["Priority band", "Objective", "Recommended work"])
ROADMAP = [
    ("Immediate operational priority", "Reduce active exposure",
     "Investigate Exchange forwarding; adjust SharePoint/OneDrive sharing defaults; restrict Temporary Access Pass scope."),
    ("High priority — identity controls", "Strengthen identity control plane",
     "Enable and operationalize PIM, reduce privileged assignments, complete authentication methods migration, validate Conditional Access coverage."),
    ("High priority — data controls", "Improve data protection",
     "Build DLP policies beyond Copilot test mode, export and validate sensitivity labels, move tuned policies to enforcement."),
    ("High priority — application governance", "Govern the SSO estate",
     "Review the 152 enterprise applications: owners, permissions, consent grants, credentials and lifecycle."),
    ("Medium-high — endpoint controls", "Validate endpoint posture",
     "Confirm Windows baseline and compliance coverage, review the Grant Local Admin policy, define mobile/BYOD strategy."),
    ("Ongoing governance", "Measure and sustain",
     "Create a monthly dashboard for privileged roles, MFA registration, DLP incidents, sharing reports, guest reviews and service health; prepare tenant-to-tenant readiness evidence."),
]
for i, r in enumerate(ROADMAP, 1):
    fill_row(t, i, list(r))

doc.add_heading("9.1 Quick Wins / Medium Term / Strategic", level=2)
t = make_table(doc, 4, 2, widths=[4.4, 11.6])
header_cells(t, ["Horizon", "Focus"])
HORIZON = [
    ("Quick wins", "Review privileged role assignments, investigate forwarding activity, tighten sharing defaults and TAP scope, configure Entra company branding."),
    ("Medium term", "Complete authentication methods migration, validate Intune compliance coverage, build the DLP policy set, establish app and guest review cycles."),
    ("Strategic improvements", "Evaluate Entra ID P2/Governance, operationalize PIM, mature information protection and AI governance, complete tenant-to-tenant readiness program."),
]
for i, kv in enumerate(HORIZON, 1):
    fill_row(t, i, list(kv))

# ---------------- 10. EVIDENCE ----------------
doc.add_heading("10. Required Follow-Up Evidence", level=1)
para(doc, "To convert this portal-based assessment into a final evidence-backed audit pack, the "
          "following exports should be collected and retained securely. These are standard senior "
          "M365 engineering artifacts and will strengthen the final presentation.")
bullets(doc, [
    "Entra role assignments and privileged role assignments export.",
    "PIM eligibility, active assignments and activation policy export (post-licensing).",
    "Conditional Access policy export with assignments, exclusions, grant and session controls.",
    "Authentication Methods policy and MFA registration status export; TAP operational restrictions.",
    "Enterprise applications export: owners, permissions, sign-in activity, credentials, certificate expiry; admin consent and OAuth grants.",
    "DLP policy export (workload scope, mode, rules, actions, incident settings); sensitivity labels and label policies.",
    "SharePoint/OneDrive tenant and site sharing settings export; external sharing and anonymous link inventory.",
    "Teams external access, guest access and policy exports; guest and external user inventory.",
    "Exchange forwarding, inbox rules, transport rules and outbound spam policy exports; SPF/DKIM/DMARC records for all production domains.",
    "Intune device inventory, compliance policies, configuration profiles, assignments and app protection policies; Grant Local Admin targeting.",
    "Microsoft Defender Secure Score, Defender for Office 365 policies and Defender for Endpoint onboarding status.",
    "Hybrid identity / Entra Connect status, Exchange hybrid configuration and MigrationWiz permissions & scope.",
])

# ---------------- 11. EXEC RECOMMENDATIONS ----------------
doc.add_heading("11. Executive Recommendations & Conclusion", level=1)
para(doc, "Kelso Industries should continue using Microsoft 365 as the strategic productivity and "
          "collaboration platform, while improving the operating controls around identity, data, "
          "devices, collaboration and service governance.")
bullets(doc, [
    "Prioritize privileged identity governance and authentication method validation — the most urgent identity issues.",
    "Conduct a cost-benefit assessment of Microsoft Entra ID P2 or Entra ID Governance for privileged identities.",
    "Validate Conditional Access before assuming policy coverage.",
    "Address external exposure: anonymous SharePoint edit links, broad Teams external access and Exchange forwarding activity.",
    "Move DLP from test/simulation mode to controlled enforcement after business validation.",
    "Review SharePoint and OneDrive permissions before expanding Copilot — Copilot inherits existing Microsoft 365 permissions.",
    "Establish service governance for Message Center, service health, roadmap changes, application ownership and agent ownership.",
    "Complete the hybrid and tenant-to-tenant readiness validations before executing the migration program.",
    "Treat tenant branding (Entra company branding / sign-in experience) as operational excellence and professional tenant experience.",
])
t = make_table(doc, 2, 1, widths=[16.0])
cell_text(t.rows[0].cells[0], "Final recommendation", bold=True, color="FFFFFF", size=11)
shade(t.rows[0].cells[0], NAVY)
cell_text(t.rows[1].cells[0],
          "Kelso Industries should approve a focused Microsoft 365 security maturity uplift program "
          "rather than a one-time configuration cleanup. The platform foundation is present and the "
          "recommended improvements are practical and achievable using Microsoft 365 capabilities "
          "already present or clearly adjacent to the current environment. The strongest investment "
          "case is Entra ID P2 / Entra ID Governance for privileged identity. Completing this program "
          "protects the business while preserving collaboration and productivity, and establishes a "
          "governed target environment for the upcoming tenant-to-tenant migrations.", size=10)

# ---------------- APPENDICES ----------------
doc.add_page_break()
doc.add_heading("Appendix A — Portal Validation Summary", level=1)
para(doc, "The following items were revalidated during the live portal review after "
          "re-authentication. Temporary local screen captures were used only for reading visible "
          "portal state and were not embedded in this report.")
t = make_table(doc, 8, 3, widths=[3.6, 6.4, 6.0])
header_cells(t, ["Area", "Validation result", "Report impact"])
VAL = [
    ("Tenant context", "Microsoft 365 and Entra portals showed kelso-industries.com.", "Confirms validation was performed against the Kelso tenant."),
    ("Conditional Access", "Policies page loaded; 4 Microsoft-managed and 4 user-created policies.", "CA remains observed; detailed export still required."),
    ("Enterprise applications", "152 applications found.", "Volume confirmed; risk focus remains ownership, consent and lifecycle."),
    ("Purview DLP", "DLP page showed 'Policy not enforced' and the Copilot DLP policy.", "DLP wording: configured in test/simulation mode, not enforced."),
    ("Microsoft Defender", "Portal opened with Defender for Business setup experience.", "Access confirmed; Secure Score and workloads require validation."),
    ("SPF / DKIM / DMARC", "Domains/DNS page could not be reliably reached during review.", "Email authentication remains 'Requires validation'."),
    ("Authentication Methods", "Page reached earlier in assessment; migration 'In progress'.", "MFA adoption/enforcement remains 'Requires validation'."),
]
for i, r in enumerate(VAL, 1):
    fill_row(t, i, list(r))

doc.add_heading("Appendix B — Reference Tenant Observations", level=1)
para(doc, "A second tenant was reviewed as a reference point for Microsoft 365 maturity "
          "patterns. It must not be treated as a direct benchmark or as proof that every "
          "configuration in that tenant is best practice; its value is to show which Microsoft "
          "capabilities matter and what a more mature pattern looks like.")
bullets(doc, [
    "Reference tenant uses Microsoft Entra ID P2, exposing PIM and identity governance capabilities — supporting the P2 investment case for Kelso's privileged population.",
    "Authentication methods migration is complete there, with Microsoft Authenticator enabled — a usable target pattern for Kelso.",
    "Copilot DLP prompt-risk telemetry was visible — demonstrating why Kelso should tune and enforce its Copilot DLP policy.",
    "Operational issues (directory sync, agents) were surfaced on the admin home page — supporting Kelso's need for service governance ownership.",
    "Caution: SMS and voice were enabled in the reference tenant; this should not be copied. Kelso should prefer stronger, phishing-resistant methods.",
])

doc.add_heading("Appendix C — Consolidation Note", level=1)
para(doc, "This report consolidates and supersedes the following working documents: the "
          "client-ready clean assessment, the reference-practices assessment, the two "
          "consultant-reviewed versions and the validated portal-review version. Wording follows "
          "the consultant-reviewed standard: observations are reported as observed portal state; "
          "items that could not be confirmed are explicitly marked 'Requires validation' or 'Not "
          "confirmed during this assessment'; the tenant ID is masked in this client-facing "
          "version; and overall risk is reported as Medium while Critical priority findings are "
          "preserved where justified.", size=10)

os.makedirs(os.path.dirname(OUT), exist_ok=True)
doc.save(OUT)
print("Saved:", OUT)
