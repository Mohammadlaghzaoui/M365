export const quickActions = [
  { label: 'New ticket', to: '/tickets', icon: 'Ticket', color: 'blue' },
  { label: 'Login / MFA issue', to: '/entra?wf=entra-cannot-login', icon: 'KeyRound', color: 'blue' },
  { label: 'Mailbox issue', to: '/exchange?wf=exo-mailbox-access', icon: 'Mail', color: 'blue' },
  { label: 'Mail flow issue', to: '/exchange?wf=exo-not-receiving', icon: 'Send', color: 'blue' },
  { label: 'Meeting room issue', to: '/exchange?wf=exo-room-disappears', icon: 'DoorOpen', color: 'orange' },
  { label: 'SharePoint access issue', to: '/sharepoint?wf=spo-no-site-access', icon: 'Boxes', color: 'blue' },
  { label: 'Teams issue', to: '/teams?wf=teams-no-access', icon: 'Users', color: 'blue' },
  { label: 'Start migration project', to: '/migration', icon: 'FolderKanban', color: 'purple' },
  { label: 'Cross-tenant migration', to: '/cross-tenant', icon: 'ArrowLeftRight', color: 'purple' },
  { label: 'BitTitan error helper', to: '/bittitan', icon: 'Cloud', color: 'purple' },
  { label: 'Syskit report helper', to: '/syskit', icon: 'Workflow', color: 'purple' },
  { label: 'Security hardening check', to: '/security', icon: 'ShieldCheck', color: 'green' },
  { label: 'Generate PowerShell', to: '/powershell', icon: 'Terminal', color: 'purple' },
  { label: 'Generate customer email', to: '/mail', icon: 'MailPlus', color: 'green' },
  { label: 'Ask AI Assistant', to: '/chat', icon: 'Bot', color: 'purple' },
];

export const dailyChecklist = [
  { id: 'dc-health', label: 'Check Microsoft 365 Service Health', link: '/monitoring' },
  { id: 'dc-signin', label: 'Check Entra ID sign-in logs', link: '/monitoring' },
  { id: 'dc-risky', label: 'Check risky users', link: '/monitoring' },
  { id: 'dc-mailflow', label: 'Check Exchange mail flow', link: '/monitoring' },
  { id: 'dc-quarantine', label: 'Check quarantine', link: '/monitoring' },
  { id: 'dc-defender', label: 'Check Defender alerts', link: '/monitoring' },
  { id: 'dc-securescore', label: 'Check Secure Score', link: '/monitoring' },
  { id: 'dc-migration', label: 'Check active migration batches', link: '/monitoring' },
  { id: 'dc-bittitan', label: 'Check BitTitan failed items', link: '/monitoring' },
  { id: 'dc-syskit', label: 'Check Syskit reports', link: '/monitoring' },
];

export const portalShortcuts = [
  { label: 'Microsoft 365 admin center', url: 'https://admin.microsoft.com' },
  { label: 'Entra admin center', url: 'https://entra.microsoft.com' },
  { label: 'Exchange admin center', url: 'https://admin.exchange.microsoft.com' },
  { label: 'SharePoint admin center', url: 'https://admin.microsoft.com/sharepoint' },
  { label: 'Teams admin center', url: 'https://admin.teams.microsoft.com' },
  { label: 'Microsoft Defender portal', url: 'https://security.microsoft.com' },
  { label: 'Microsoft Purview portal', url: 'https://purview.microsoft.com' },
  { label: 'Intune admin center', url: 'https://intune.microsoft.com' },
  { label: 'Service Health', url: 'https://admin.microsoft.com/Adminportal/Home#/servicehealth' },
  { label: 'BitTitan MigrationWiz', url: 'https://migrationwiz.bittitan.com' },
  { label: 'Microsoft Graph Explorer', url: 'https://developer.microsoft.com/graph/graph-explorer' },
  { label: 'M365 network status', url: 'https://status.cloud.microsoft' },
];

export const favoritePsDefaults = [
  'ps-exo-trace',
  'ps-exo-perms',
  'ps-exo-room-get',
  'ps-entra-getuser',
  'ps-ct-getmiguser',
];
