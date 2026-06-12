# -*- coding: utf-8 -*-
"""
Generates a professional M365 tenant-to-tenant migration workbook per company.
Source tenant: the company itself | Target tenant: Kelso
"""
import os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule

COMPANIES = [
    "RGD", "Retrofit", "Quality H&A", "RMR", "LLR", "YES", "Frey Lutz",
    "Icon", "Wallace", "Markade", "Fab-Logix", "Premier", "Strictly",
]
TARGET = "Kelso"
OUT_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "Migration-Excels")

# ---------- Style definitions ----------
C_DARK = "1F3864"      # dark blue
C_MED = "2F5597"       # medium blue
C_LIGHT = "D6E4F0"     # light blue
C_ACCENT = "ED7D31"    # orange accent
C_GREY = "F2F2F2"

F_TITLE = Font(name="Calibri", size=20, bold=True, color="FFFFFF")
F_SUB = Font(name="Calibri", size=12, bold=False, color="FFFFFF")
F_HDR = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
F_PHASE = Font(name="Calibri", size=12, bold=True, color="FFFFFF")
F_BODY = Font(name="Calibri", size=11)
F_BOLD = Font(name="Calibri", size=11, bold=True)

FILL_DARK = PatternFill("solid", fgColor=C_DARK)
FILL_MED = PatternFill("solid", fgColor=C_MED)
FILL_LIGHT = PatternFill("solid", fgColor=C_LIGHT)
FILL_ACCENT = PatternFill("solid", fgColor=C_ACCENT)
FILL_GREY = PatternFill("solid", fgColor=C_GREY)

THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

WRAP = Alignment(wrap_text=True, vertical="top")
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)

STATUS_LIST = '"Not started,In progress,Completed,Blocked,N/A"'
PRIO_LIST = '"High,Medium,Low"'

STATUS_COLORS = {
    "Completed": "C6EFCE",
    "In progress": "FFEB9C",
    "Blocked": "FFC7CE",
}


def status_validation(ws, col, first_row, last_row):
    dv = DataValidation(type="list", formula1=STATUS_LIST, allow_blank=True)
    dv.error = "Choose a value from the list"
    dv.errorTitle = "Invalid status"
    ws.add_data_validation(dv)
    dv.add(f"{col}{first_row}:{col}{last_row}")
    for value, color in STATUS_COLORS.items():
        ws.conditional_formatting.add(
            f"{col}{first_row}:{col}{last_row}",
            CellIsRule(operator="equal", formula=[f'"{value}"'],
                       fill=PatternFill("solid", fgColor=color)))


def prio_validation(ws, col, first_row, last_row):
    dv = DataValidation(type="list", formula1=PRIO_LIST, allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(f"{col}{first_row}:{col}{last_row}")
    ws.conditional_formatting.add(
        f"{col}{first_row}:{col}{last_row}",
        CellIsRule(operator="equal", formula=['"High"'],
                   fill=PatternFill("solid", fgColor="FFC7CE")))


def banner(ws, company, subtitle, ncols):
    last = get_column_letter(ncols)
    ws.merge_cells(f"A1:{last}1")
    ws.merge_cells(f"A2:{last}2")
    c = ws["A1"]
    c.value = f"M365 Tenant-to-Tenant Migration  |  {company}  →  {TARGET}"
    c.font = F_TITLE
    c.fill = FILL_DARK
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    s = ws["A2"]
    s.value = subtitle
    s.font = F_SUB
    s.fill = FILL_MED
    s.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[1].height = 32
    ws.row_dimensions[2].height = 20
    for col in range(1, ncols + 1):
        ws.cell(row=1, column=col).fill = FILL_DARK
        ws.cell(row=2, column=col).fill = FILL_MED


def header_row(ws, row, headers, widths):
    for i, (h, w) in enumerate(zip(headers, widths), start=1):
        cell = ws.cell(row=row, column=i, value=h)
        cell.font = F_HDR
        cell.fill = FILL_MED
        cell.alignment = CENTER
        cell.border = BORDER
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.row_dimensions[row].height = 22


def task_sheet(wb, company, title, tab_color, intro, phases):
    """phases: list of (phase name, [(task, description, default owner), ...])"""
    ws = wb.create_sheet(title)
    ws.sheet_properties.tabColor = tab_color
    headers = ["No", "Task", "Description / Instructions", "Owner",
               "Priority", "Status", "Due date", "Completed on", "Notes"]
    widths = [6, 38, 70, 20, 11, 13, 13, 13, 35]
    banner(ws, company, intro, len(headers))
    header_row(ws, 4, headers, widths)

    row = 5
    nr = 1
    for phase_name, tasks in phases:
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=len(headers))
        pc = ws.cell(row=row, column=1, value=phase_name)
        pc.font = F_PHASE
        pc.fill = FILL_ACCENT
        pc.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        for col in range(1, len(headers) + 1):
            ws.cell(row=row, column=col).fill = FILL_ACCENT
        ws.row_dimensions[row].height = 20
        row += 1
        for task, desc, owner in tasks:
            values = [nr, task, desc, owner, "", "Not started", "", "", ""]
            for col, v in enumerate(values, start=1):
                cell = ws.cell(row=row, column=col, value=v)
                cell.font = F_BODY
                cell.border = BORDER
                cell.alignment = WRAP if col in (2, 3, 9) else CENTER
            if row % 2 == 0:
                for col in range(1, len(headers) + 1):
                    if ws.cell(row=row, column=col).fill.fgColor.rgb in (None, "00000000"):
                        ws.cell(row=row, column=col).fill = FILL_GREY
            nr += 1
            row += 1

    last_row = row - 1
    status_validation(ws, "F", 5, last_row)
    prio_validation(ws, "E", 5, last_row)
    ws.auto_filter.ref = f"A4:I{last_row}"
    ws.freeze_panes = "A5"
    return ws


# ---------- Phase content ----------

def phases_preparation(company):
    return [
        ("PHASE 1 — PROJECT & GOVERNANCE", [
            ("Kick-off & define scope",
             f"Define project scope: which workloads of {company} migrate to {TARGET} (Exchange, OneDrive, SharePoint, Teams, apps). Identify stakeholders and decision makers.",
             "Project lead"),
            ("Migration planning & timeline",
             "Create a detailed plan with milestones, migration waves and the cutover date. Take holidays and business-critical periods into account.",
             "Project lead"),
            ("Create communication plan",
             "Inform end users in time: what changes, when, and what they need to do themselves (new password, re-register MFA, Outlook profile, Teams).",
             "Project lead"),
            ("Define rollback / contingency plan",
             "Describe a fallback scenario per workload in case the migration fails (revert MX, forwarding, restore from backup).",
             "M365 Engineer"),
            ("Select migration tool & licenses",
             "Select a tool and purchase licenses (e.g. BitTitan MigrationWiz, Quest On Demand, AvePoint Fly, ShareGate). Determine number of mailboxes/users for license count.",
             "M365 Engineer"),
        ]),
        ("PHASE 2 — ACCESS & ACCOUNTS", [
            ("Global Admin access source tenant",
             f"Obtain/create a Global Administrator account in the {company} tenant (incl. MFA). Document a break-glass account.",
             "M365 Engineer"),
            (f"Global Admin access target tenant ({TARGET})",
             f"Global Administrator account in the {TARGET} tenant. Verify sufficient rights for user creation, Exchange, SharePoint and Teams admin.",
             "M365 Engineer"),
            ("Service accounts for migration tool",
             "Create dedicated migration admin accounts in source and target (without MFA blocking the tool, or with app registration/modern auth). Remove afterwards.",
             "M365 Engineer"),
            ("App registrations & API permissions",
             "App registration in Entra ID for the migration tool with the correct Graph/EWS permissions; grant admin consent in both tenants.",
             "M365 Engineer"),
        ]),
        ("PHASE 3 — LICENSES & TARGET TENANT CAPACITY", [
            (f"Verify licenses in {TARGET}",
             f"Sufficient M365 licenses (Business Premium/E3/E5 etc.) available in {TARGET} for all {company} users to be migrated.",
             "Project lead"),
            ("Check SharePoint storage capacity",
             f"Verify that the SharePoint storage quota in {TARGET} can hold the {company} data (use the inventory figures).",
             "M365 Engineer"),
            ("Define naming convention",
             f"Define the UPN/email format in {TARGET} (e.g. firstname.lastname@kelso-domain) and identify conflicts with existing {TARGET} users.",
             "M365 Engineer"),
        ]),
        ("PHASE 4 — DOMAIN & DNS PREPARATION", [
            ("Arrange DNS management access",
             f"Verify access to the DNS registrar/zone of the {company} domain. Lower the TTL of MX/Autodiscover records to 300-3600 sec before cutover.",
             "M365 Engineer"),
            ("Prepare domain move",
             f"Plan for removing the mail domain from the {company} tenant and adding + verifying it in {TARGET} (prepare the verification TXT record). Note: the domain can only move once no objects use it anymore.",
             "M365 Engineer"),
            ("External DNS dependencies",
             "Inventory and reconfigure SPF, DKIM, DMARC, MTA-STS and any 3rd-party records (printers/scanners, applications using SMTP relay).",
             "M365 Engineer"),
        ]),
        ("PHASE 5 — END USERS & WORKPLACE", [
            ("Inventory workstations",
             "Overview of all PCs/laptops: Entra-joined or hybrid? Intune-managed? Plan for moving devices to the Kelso tenant.",
             "Workplace admin"),
            ("Outlook/OneDrive profile plan",
             "Create a procedure for a new Outlook profile and OneDrive re-link on every workstation after cutover (manual or scripted).",
             "Workplace admin"),
            ("Mobile devices",
             "Plan for re-enrolling phones/tablets (Outlook mobile, Teams, Company Portal, re-register MFA/Authenticator).",
             "Workplace admin"),
            ("Password & MFA instructions",
             "Instruction document for users: first sign-in to the Kelso tenant, set password, register MFA.",
             "Service desk"),
        ]),
    ]


def phases_inventory(company):
    return [
        ("INVENTORY — IDENTITY", [
            ("Export user overview",
             f"Export all users from {company} (Entra ID > Users or Graph/PowerShell: Get-MgUser). Record: UPN, display name, licenses, last sign-in, active/inactive.",
             "M365 Engineer"),
            ("Shared mailboxes & resources",
             "Export shared mailboxes, room/equipment mailboxes and their delegations (Full Access, Send As, Send on Behalf) (Get-Mailbox / Get-MailboxPermission).",
             "M365 Engineer"),
            ("Distribution lists & M365 groups",
             "Export all distribution groups, mail-enabled security groups and Microsoft 365 groups incl. members and owners.",
             "M365 Engineer"),
            ("Guest accounts (B2B)",
             "Inventory external guests in the tenant; decide which ones need to be re-invited in Kelso.",
             "M365 Engineer"),
            ("Document admin roles",
             "Who holds which admin role in the source tenant; determine which roles are needed in Kelso.",
             "M365 Engineer"),
        ]),
        ("INVENTORY — EXCHANGE ONLINE", [
            ("Mailbox sizes & item counts",
             "Export size and item count per mailbox (Get-MailboxStatistics) — determines migration duration and possible pre-staging.",
             "M365 Engineer"),
            ("Archive mailboxes",
             "Identify In-Place Archives (size; auto-expanding archives require extra attention/tooling).",
             "M365 Engineer"),
            ("Mail flow rules & connectors",
             "Document transport rules, connectors, anti-spam/anti-phish policies and email disclaimers to rebuild in Kelso.",
             "M365 Engineer"),
            ("Forwarding & inbox rules",
             "Inventory mailbox forwarding and critical inbox rules (they do not always migrate automatically).",
             "M365 Engineer"),
            ("Aliases & proxy addresses",
             "Export all SMTP aliases per mailbox so they can be recreated in Kelso.",
             "M365 Engineer"),
            ("Litigation hold & retention",
             "Identify mailboxes with Litigation Hold/retention policies — secure the data before migration (hold data does not migrate by itself).",
             "M365 Engineer"),
        ]),
        ("INVENTORY — ONEDRIVE & SHAREPOINT", [
            ("OneDrive overview",
             "Export OneDrive URL, size and file count per user (SharePoint admin center or PowerShell).",
             "M365 Engineer"),
            ("Inventory SharePoint sites",
             "All sites (team/communication), size, owners, last activity. Flag obsolete sites NOT to migrate (cleaning up = faster migration).",
             "M365 Engineer"),
            ("Permissions & external sharing links",
             "Document unique permissions, external sharing and anonymous links — external links break after migration.",
             "M365 Engineer"),
            ("Identify obsolete data",
             f"Agree with {company} which data can be archived or deleted before the migration.",
             "Project lead"),
        ]),
        ("INVENTORY — TEAMS", [
            ("Export teams & channels",
             "Export all teams, channels (incl. private/shared channels!), members and owners. Private channels have their own SharePoint sites.",
             "M365 Engineer"),
            ("Assess chat history",
             "Decide whether 1:1/group chats must be migrated (slow and only partially possible; often only team channel messages are migrated).",
             "Project lead"),
            ("Apps, tabs & connectors",
             "Document installed Teams apps, tabs (Planner, OneNote, websites) and webhooks per team — these must be redone manually.",
             "M365 Engineer"),
            ("Telephony/voice (if present)",
             "Inventory Teams Phone numbers, call queues and auto attendants; plan number porting.",
             "M365 Engineer"),
        ]),
        ("INVENTORY — APPLICATIONS & OTHER", [
            ("Enterprise apps & SSO",
             "Applications that sign in via Entra SSO on the source tenant (SaaS apps, VPN, etc.) — must be reconnected to Kelso.",
             "M365 Engineer"),
            ("Power Platform",
             "Inventory Power Automate flows, Power Apps and Power BI reports/workspaces — these rarely migrate automatically.",
             "M365 Engineer"),
            ("Intune & compliance policies",
             "Document device configuration, compliance policies, conditional access and app protection policies to rebuild in Kelso.",
             "M365 Engineer"),
            ("SMTP relay & multifunction devices",
             "Printers, scanners and applications that send mail via the source tenant — reconfigure to Kelso.",
             "Workplace admin"),
            ("Complete risk analysis",
             "Summarize all findings on the 'Risks & Issues' tab and discuss with the customer.",
             "Project lead"),
        ]),
    ]


def phases_identity(company):
    return [
        ("CREATE USERS IN KELSO", [
            ("Create user mapping table",
             "Source UPN → target UPN mapping for all users ('User Mapping' tab). This is the basis for the migration tool.",
             "M365 Engineer"),
            ("Provision users in Kelso",
             "Create accounts in the Kelso tenant (CSV import, Graph or scripted). Temporarily use @kelso-onmicrosoft or the target domain as UPN.",
             "M365 Engineer"),
            ("Assign licenses",
             "Assign the correct licenses per user so the mailbox/OneDrive gets provisioned (speed up OneDrive pre-provisioning via PowerShell).",
             "M365 Engineer"),
            ("Create groups & DLs",
             "Rebuild distribution lists, security groups and M365 groups in Kelso with the correct members/owners.",
             "M365 Engineer"),
            ("Create shared & resource mailboxes",
             "Create shared mailboxes and room/equipment mailboxes with the same delegations as in the source.",
             "M365 Engineer"),
        ]),
        ("SECURITY & POLICY", [
            ("Prepare MFA/Conditional Access",
             "Review CA policies in Kelso; make sure migration service accounts are excluded during the migration.",
             "M365 Engineer"),
            ("Password policy & SSPR",
             "Distribute initial passwords via a secure channel; configure Self-Service Password Reset.",
             "M365 Engineer"),
            ("Assign admin roles",
             "Assign required admin roles in Kelso following least privilege.",
             "M365 Engineer"),
        ]),
    ]


def phases_exchange(company):
    return [
        ("PRE-STAGE (BEFORE CUTOVER)", [
            ("Configure migration tool",
             f"Connect source ({company}) and target ({TARGET}) endpoints in the migration tool; import the mapping table; run a test batch with 2-3 mailboxes.",
             "M365 Engineer"),
            ("Pre-stage sync mailboxes",
             "First full sync of all mailbox data (mail, calendar, contacts, tasks) while users still work on the source. Start the largest mailboxes first.",
             "M365 Engineer"),
            ("Migrate archives",
             "Migrate In-Place Archives separately to archives in Kelso (check tool support).",
             "M365 Engineer"),
            ("Run delta syncs",
             "Daily delta syncs until cutover so the delta on cutover day is minimal.",
             "M365 Engineer"),
            ("Review error reports",
             "Review errors/skipped items per batch and resolve them (corrupt items, items >150MB, throttling).",
             "M365 Engineer"),
        ]),
        ("CUTOVER DAY", [
            ("Freeze mail flow / final delta",
             "Run the final delta sync. Optional: block sending on source mailboxes during the cutover window.",
             "M365 Engineer"),
            ("Switch MX records",
             f"Switch MX, Autodiscover, SPF, DKIM and DMARC to the {TARGET} tenant. Prepare and enable DKIM in Kelso beforehand.",
             "M365 Engineer"),
            ("Move domain (if applicable)",
             f"Detach the mail domain from the {company} tenant (first switch all aliases/UPNs to onmicrosoft) and add + verify it in {TARGET}; then switch UPNs and primary SMTP.",
             "M365 Engineer"),
            ("Test mail flow",
             "Test inbound/outbound mail (internal, external, to shared mailboxes, distribution lists). Check mail tips and signatures.",
             "M365 Engineer"),
            ("Forwarding source → target (safety net)",
             "Set up temporary forwarding on source mailboxes for mail that still arrives via the old route.",
             "M365 Engineer"),
        ]),
        ("AFTER CUTOVER", [
            ("Post-cutover delta",
             "Run a final delta to catch mail delivered to the source during cutover.",
             "M365 Engineer"),
            ("Renew Outlook profiles",
             "New Outlook profiles on all workstations; rebuild cache; restore signatures.",
             "Service desk"),
            ("Reconfigure mobile mail",
             "Reconnect Outlook mobile to the Kelso account.",
             "Service desk"),
        ]),
    ]


def phases_files(company):
    return [
        ("ONEDRIVE", [
            ("OneDrive pre-provisioning",
             "Pre-create OneDrives in Kelso (Request-SPOPersonalSite) so the migration tool can write immediately.",
             "M365 Engineer"),
            ("OneDrive pre-stage sync",
             "Full first sync of all OneDrive data; then delta syncs until cutover.",
             "M365 Engineer"),
            ("Version history & metadata",
             "Check tool settings: include versions, authors and timestamps where possible.",
             "M365 Engineer"),
            ("Communicate shared links",
             "Inform users that existing sharing links will break and items must be re-shared.",
             "Service desk"),
            ("Re-link OneDrive client",
             "After cutover, unlink OneDrive from the source and link it to the Kelso account on every workstation (avoid sync conflicts: rename the old folder).",
             "Workplace admin"),
        ]),
        ("SHAREPOINT", [
            ("Create site structure in Kelso",
             "Create target sites (or let the tool create them) following the agreed structure; set site owners.",
             "M365 Engineer"),
            ("SharePoint pre-stage sync",
             "Full sync of all sites to migrate incl. document libraries, lists and permissions; deltas until cutover.",
             "M365 Engineer"),
            ("Validate permissions",
             "Spot check per site: are the permissions correct (owners/members/visitors, unique permissions)?",
             "M365 Engineer"),
            ("Shortcuts & synced libraries",
             "Users with synced SharePoint libraries: remove the old sync, set up a new sync to the Kelso sites.",
             "Workplace admin"),
            ("Final delta & verification",
             "Final delta at cutover; compare item counts source vs. target; report discrepancies.",
             "M365 Engineer"),
        ]),
    ]


def phases_teams(company):
    return [
        ("TEAMS — STRUCTURE & DATA", [
            ("Create teams in Kelso",
             "Rebuild teams + channels (incl. private/shared channels) or let the migration tool create them; link owners and members via the mapping table.",
             "M365 Engineer"),
            ("Migrate channel files",
             "Migrate the files of every channel (SharePoint behind the team); don't forget private channel sites.",
             "M365 Engineer"),
            ("Migrate channel messages",
             "Migrate channel conversations if desired (often as an HTML archive or via a tool with limitations). Manage user expectations.",
             "M365 Engineer"),
            ("Chats (1:1/group)",
             "If in scope: run chat migration (limited; alternative = readable archive). Document the decision.",
             "Project lead"),
            ("Restore tabs, apps & connections",
             "Manually restore Planner, OneNote, website tabs and app connections per team using the inventory.",
             "M365 Engineer"),
        ]),
        ("TEAMS — AFTERCARE", [
            ("Reset Teams client",
             "Sign users out/clear cache and sign in with the Kelso account; recreate calendar meetings if Teams links point to the source tenant.",
             "Service desk"),
            ("Recurring meetings",
             "Recurring Teams meetings contain links to the old tenant — have organizers recreate them after cutover.",
             "Service desk"),
            ("Move telephony (if present)",
             "Port/move numbers, rebuild call queues and auto attendants in Kelso.",
             "M365 Engineer"),
        ]),
    ]


def phases_cutover(company):
    return [
        ("CUTOVER CHECKLIST (GO/NO-GO)", [
            ("Go/No-Go decision",
             "Walk through the checklist with the project team: pre-stage syncs ≥95% complete, no blocking errors, communication sent, support ready.",
             "Project lead"),
            ("Send end-user communication",
             "Final instruction email: timing, what to do Monday morning, where to get help.",
             "Project lead"),
            ("Final delta all workloads",
             "Final delta sync for mail, OneDrive, SharePoint and Teams.",
             "M365 Engineer"),
            ("Execute DNS cutover",
             "Switch MX/Autodiscover/SPF/DKIM/DMARC (see Exchange tab).",
             "M365 Engineer"),
            ("Switch UPN & primary SMTP",
             "Switch users in Kelso to the production domain as primary address.",
             "M365 Engineer"),
            ("Restrict source tenant",
             "Block sign-in on source accounts (after validation) to prevent confusion and double usage.",
             "M365 Engineer"),
        ]),
        ("VALIDATION AFTER CUTOVER", [
            ("Mail flow validation",
             "Test internal/external/inbound/outbound, shared mailboxes, DLs, signatures, calendar delegations.",
             "M365 Engineer"),
            ("Data spot checks",
             "Spot check per user group on mail, OneDrive, SharePoint and Teams files (counts and recent items).",
             "M365 Engineer"),
            ("Workstations switched",
             "All PCs: new profile, OneDrive linked, Teams signed in, printers/scan-to-mail working.",
             "Workplace admin"),
            ("Start hypercare period",
             "1-2 weeks of elevated support; log issues on the 'Risks & Issues' tab.",
             "Service desk"),
        ]),
        ("CLOSURE & DECOMMISSION", [
            ("Deliver migration report",
             f"Final report: migrated counts, open items, deviations, lessons learned for {company}.",
             "Project lead"),
            ("Clean up forwarding & service accounts",
             "Remove temporary forwarding, migration accounts and app registrations in both tenants.",
             "M365 Engineer"),
            ("Decommission source tenant",
             "After the agreed retention period: cancel licenses, secure a data export/backup, retire or park the tenant.",
             "M365 Engineer"),
            ("Documentation & handover",
             "Update the Kelso environment's admin documentation and hand it over to operations.",
             "Project lead"),
        ]),
    ]


# ---------- Special tabs ----------

def overview_sheet(wb, company):
    ws = wb.create_sheet("Overview", 0)
    ws.sheet_properties.tabColor = C_DARK
    banner(ws, company, "Project overview & progress per phase", 6)

    info = [
        ("Source tenant", company),
        ("Target tenant", TARGET),
        ("Project lead", ""),
        ("M365 Engineer", ""),
        ("Migration tool", ""),
        ("Planned cutover date", ""),
        ("Number of users", ""),
        ("Number of mailboxes (incl. shared)", ""),
        ("Total data size (GB)", ""),
        ("Project status", "Not started"),
    ]
    row = 4
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=6)
    c = ws.cell(row=row, column=1, value="PROJECT DETAILS")
    c.font = F_PHASE; c.fill = FILL_ACCENT
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    for col in range(1, 7):
        ws.cell(row=row, column=col).fill = FILL_ACCENT
    row += 1
    for label, val in info:
        lc = ws.cell(row=row, column=1, value=label)
        lc.font = F_BOLD; lc.fill = FILL_LIGHT; lc.border = BORDER
        ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=3)
        vc = ws.cell(row=row, column=2, value=val)
        vc.font = F_BODY; vc.border = BORDER
        ws.cell(row=row, column=3).border = BORDER
        row += 1

    row += 1
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=6)
    c = ws.cell(row=row, column=1, value="PROGRESS PER PHASE")
    c.font = F_PHASE; c.fill = FILL_ACCENT
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    for col in range(1, 7):
        ws.cell(row=row, column=col).fill = FILL_ACCENT
    row += 1
    hdrs = ["Phase", "Tab", "Status", "Start date", "End date", "Remarks"]
    for i, h in enumerate(hdrs, 1):
        c = ws.cell(row=row, column=i, value=h)
        c.font = F_HDR; c.fill = FILL_MED; c.alignment = CENTER; c.border = BORDER
    row += 1
    first_status = row
    phases = [
        ("1. Preparation", "1. Preparation"),
        ("2. Analysis & Inventory", "2. Analysis & Inventory"),
        ("3. Identity & Users", "3. Identity"),
        ("4. Exchange Online", "4. Exchange"),
        ("5. OneDrive & SharePoint", "5. OneDrive-SharePoint"),
        ("6. Teams", "6. Teams"),
        ("7. Cutover & Aftercare", "7. Cutover & Aftercare"),
    ]
    for name, tab in phases:
        ws.cell(row=row, column=1, value=name).font = F_BOLD
        ws.cell(row=row, column=2, value=tab).font = F_BODY
        ws.cell(row=row, column=3, value="Not started").alignment = CENTER
        for col in range(1, 7):
            ws.cell(row=row, column=col).border = BORDER
        row += 1
    status_validation(ws, "C", first_status, row - 1)

    widths = [32, 26, 16, 14, 14, 40]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    return ws


def mapping_sheet(wb, company):
    ws = wb.create_sheet("User Mapping")
    ws.sheet_properties.tabColor = "7030A0"
    headers = ["No", "Display name", f"Source UPN ({company})", f"Target UPN ({TARGET})",
               "Target license", "Mailbox type", "Mailbox size (GB)", "OneDrive (GB)",
               "Migrated?", "Notes"]
    widths = [6, 26, 34, 34, 18, 16, 16, 14, 13, 30]
    banner(ws, company, "User mapping source → target (fill in from the inventory export)", len(headers))
    header_row(ws, 4, headers, widths)
    for r in range(5, 55):
        ws.cell(row=r, column=1, value=r - 4).alignment = CENTER
        for col in range(1, len(headers) + 1):
            ws.cell(row=r, column=col).border = BORDER
            ws.cell(row=r, column=col).font = F_BODY
    dv = DataValidation(type="list", formula1='"Yes,No,In progress"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add("I5:I54")
    dv2 = DataValidation(type="list", formula1='"User,Shared,Room,Equipment"', allow_blank=True)
    ws.add_data_validation(dv2)
    dv2.add("F5:F54")
    ws.freeze_panes = "A5"
    ws.auto_filter.ref = "A4:J54"
    return ws


def risk_sheet(wb, company):
    ws = wb.create_sheet("Risks & Issues")
    ws.sheet_properties.tabColor = "C00000"
    headers = ["No", "Date", "Type", "Description", "Impact", "Priority",
               "Owner", "Status", "Resolution / Mitigation"]
    widths = [6, 12, 12, 50, 30, 11, 18, 13, 45]
    banner(ws, company, "Risk and issue log — track all findings and incidents here", len(headers))
    header_row(ws, 4, headers, widths)
    seed = [
        ("Risk", "External sharing links (SharePoint/OneDrive) stop working after migration.",
         "Users/external parties lose access to shared files.",
         "Communicate; re-share the most important shares after migration."),
        ("Risk", "Recurring Teams meeting links point to the old tenant.",
         "Meetings not accessible after cutover.",
         "Have organizers recreate them after cutover."),
        ("Risk", "Microsoft throttling during large data migrations.",
         "Migration takes longer than planned.",
         "Start pre-staging early; spread migration waves."),
        ("Risk", "Items >150 MB or corrupt items are skipped.",
         "Missing files/mail items.",
         "Review error reports per batch; handle manually afterwards."),
        ("Risk", "Domain cannot move due to remaining objects.",
         "Cutover delay.",
         "Switch all aliases/UPNs to .onmicrosoft.com beforehand and validate."),
    ]
    row = 5
    for i, (typ, desc, imp, mit) in enumerate(seed, 1):
        vals = [i, "", typ, desc, imp, "High", "", "Open", mit]
        for col, v in enumerate(vals, 1):
            c = ws.cell(row=row, column=col, value=v)
            c.font = F_BODY; c.border = BORDER
            c.alignment = WRAP if col in (4, 5, 9) else CENTER
        row += 1
    for r in range(row, row + 20):
        for col in range(1, len(headers) + 1):
            ws.cell(row=r, column=col).border = BORDER
    last = row + 19
    dv = DataValidation(type="list", formula1='"Open,In progress,Resolved,Accepted"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(f"H5:H{last}")
    dvt = DataValidation(type="list", formula1='"Risk,Issue,Decision,Action"', allow_blank=True)
    ws.add_data_validation(dvt)
    dvt.add(f"C5:C{last}")
    prio_validation(ws, "F", 5, last)
    ws.conditional_formatting.add(
        f"H5:H{last}", CellIsRule(operator="equal", formula=['"Resolved"'],
                                  fill=PatternFill("solid", fgColor="C6EFCE")))
    ws.conditional_formatting.add(
        f"H5:H{last}", CellIsRule(operator="equal", formula=['"Open"'],
                                  fill=PatternFill("solid", fgColor="FFC7CE")))
    ws.freeze_panes = "A5"
    ws.auto_filter.ref = f"A4:I{last}"
    return ws


def build_workbook(company):
    wb = Workbook()
    wb.remove(wb.active)
    overview_sheet(wb, company)
    task_sheet(wb, company, "1. Preparation", "4472C4",
               "Phase 1 — Preparation: governance, access, licenses, DNS and workplace",
               phases_preparation(company))
    task_sheet(wb, company, "2. Analysis & Inventory", "70AD47",
               "Phase 2 — Analysis & Inventory: map out the complete source tenant",
               phases_inventory(company))
    mapping_sheet(wb, company)
    task_sheet(wb, company, "3. Identity", "FFC000",
               "Phase 3 — Provision identity & users in the Kelso tenant",
               phases_identity(company))
    task_sheet(wb, company, "4. Exchange", "ED7D31",
               "Phase 4 — Exchange Online migration: pre-stage, cutover and aftercare",
               phases_exchange(company))
    task_sheet(wb, company, "5. OneDrive-SharePoint", "5B9BD5",
               "Phase 5 — OneDrive & SharePoint data migration",
               phases_files(company))
    task_sheet(wb, company, "6. Teams", "7030A0",
               "Phase 6 — Microsoft Teams migration",
               phases_teams(company))
    task_sheet(wb, company, "7. Cutover & Aftercare", "C00000",
               "Phase 7 — Cutover, validation, hypercare and decommission of the source tenant",
               phases_cutover(company))
    risk_sheet(wb, company)
    return wb


def safe_name(name):
    for ch in '\\/:*?"<>|&':
        name = name.replace(ch, "-")
    return name.replace(" ", "_")


if __name__ == "__main__":
    os.makedirs(OUT_DIR, exist_ok=True)
    for comp in COMPANIES:
        wb = build_workbook(comp)
        path = os.path.join(OUT_DIR, f"M365_Migration_{safe_name(comp)}_to_{TARGET}.xlsx")
        wb.save(path)
        print("OK", path)
    print("Done:", len(COMPANIES), "workbooks in", OUT_DIR)
