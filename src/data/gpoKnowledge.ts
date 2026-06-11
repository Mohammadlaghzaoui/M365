// GPO → Intune knowledge base. Real, expandable mappings used by the GPO Advisor
// to match parsed policy settings to their modern Intune/M365 equivalent.

export type GpoTarget =
  | 'Settings Catalog' | 'Administrative Templates' | 'Security Baseline'
  | 'Endpoint Security' | 'PowerShell Script' | 'Win32 App'
  | 'OneDrive Policy' | 'Defender Policy' | 'Unsupported' | 'Manual';

export type GpoSupport = 'Supported' | 'ScriptRequired' | 'Manual' | 'Unsupported';

export interface GpoMapping {
  id: string;
  /** lower-cased keywords matched against the policy name/registry path */
  match: string[];
  gpoName: string;
  scope: 'Computer' | 'User' | 'Both';
  category: string;
  target: GpoTarget;
  intunePath: string;
  setting: string;
  support: GpoSupport;
  confidence: number; // 0-100
  notes?: string;
  custom?: boolean;
}

export const seedGpoMappings: GpoMapping[] = [
  { id: 'gpo-cmd', match: ['disallowcmd', 'command prompt', 'prevent access to the command prompt'], gpoName: 'Prevent access to the command prompt', scope: 'User', category: 'System', target: 'Settings Catalog', intunePath: 'User Configuration > Administrative Templates > System', setting: 'Prevent access to the command prompt', support: 'Supported', confidence: 98 },
  { id: 'gpo-regedit', match: ['disableregistrytools', 'registry editing tools'], gpoName: 'Prevent access to registry editing tools', scope: 'User', category: 'System', target: 'Settings Catalog', intunePath: 'User Configuration > Administrative Templates > System', setting: 'Prevent access to registry editing tools', support: 'Supported', confidence: 97 },
  { id: 'gpo-controlpanel', match: ['nocontrolpanel', 'prohibit access to control panel'], gpoName: 'Prohibit access to Control Panel and PC settings', scope: 'User', category: 'Control Panel', target: 'Settings Catalog', intunePath: 'User Configuration > Administrative Templates > Control Panel', setting: 'Prohibit access to Control Panel and PC settings', support: 'Supported', confidence: 96 },
  { id: 'gpo-screensaver', match: ['screensavetimeout', 'screensaveractive', 'screen saver timeout'], gpoName: 'Screen saver timeout', scope: 'User', category: 'Personalization', target: 'Settings Catalog', intunePath: 'User Configuration > Administrative Templates > Control Panel > Personalization', setting: 'Screen saver timeout', support: 'Supported', confidence: 95 },
  { id: 'gpo-password-length', match: ['minimumpasswordlength', 'minimum password length'], gpoName: 'Minimum password length', scope: 'Computer', category: 'Account Policies', target: 'Security Baseline', intunePath: 'Endpoint security > Account protection / Settings Catalog > Password Policy', setting: 'Minimum password length', support: 'Supported', confidence: 94, notes: 'Cloud-only devices: password complexity largely handled by Entra; baseline still applies locally.' },
  { id: 'gpo-password-complexity', match: ['passwordcomplexity', 'password must meet complexity'], gpoName: 'Password must meet complexity requirements', scope: 'Computer', category: 'Account Policies', target: 'Security Baseline', intunePath: 'Endpoint security > Account protection', setting: 'Password complexity', support: 'Supported', confidence: 93 },
  { id: 'gpo-lockout', match: ['lockoutthreshold', 'account lockout threshold'], gpoName: 'Account lockout threshold', scope: 'Computer', category: 'Account Policies', target: 'Settings Catalog', intunePath: 'Settings Catalog > Account Lockout', setting: 'Account lockout threshold', support: 'Supported', confidence: 92 },
  { id: 'gpo-firewall', match: ['windowsfirewall', 'enablefirewall', 'firewall: protect all network'], gpoName: 'Windows Firewall: Protect all network connections', scope: 'Computer', category: 'Windows Firewall', target: 'Endpoint Security', intunePath: 'Endpoint security > Firewall', setting: 'Enable Windows Defender Firewall', support: 'Supported', confidence: 95 },
  { id: 'gpo-defender-rtp', match: ['disableantispyware', 'disablerealtimemonitoring', 'real-time protection'], gpoName: 'Turn off real-time protection', scope: 'Computer', category: 'Microsoft Defender Antivirus', target: 'Defender Policy', intunePath: 'Endpoint security > Antivirus (Microsoft Defender)', setting: 'Real-time protection', support: 'Supported', confidence: 96, notes: 'Invert intent: configure Defender AV policy in Endpoint security rather than disabling.' },
  { id: 'gpo-bitlocker', match: ['bitlocker', 'fdvencryptiontype', 'fixed data drives'], gpoName: 'BitLocker drive encryption settings', scope: 'Computer', category: 'BitLocker', target: 'Endpoint Security', intunePath: 'Endpoint security > Disk encryption', setting: 'BitLocker policy', support: 'Supported', confidence: 90 },
  { id: 'gpo-folder-redirection', match: ['folderredirection', 'redirect', 'documents folder redirect'], gpoName: 'Folder Redirection (Documents/Desktop)', scope: 'User', category: 'Folder Redirection', target: 'OneDrive Policy', intunePath: 'Settings Catalog > OneDrive > Known Folder Move (KFM)', setting: 'Silently move Windows known folders to OneDrive', support: 'Supported', confidence: 88, notes: 'Replace classic folder redirection with OneDrive Known Folder Move — do not lift-and-shift UNC redirection.' },
  { id: 'gpo-drive-map', match: ['drivemapsettings', 'drive map', 'mapped drive', 'net use'], gpoName: 'Drive Maps (Group Policy Preferences)', scope: 'User', category: 'Preferences - Drive Maps', target: 'PowerShell Script', intunePath: 'Devices > Scripts and remediations (PowerShell)', setting: 'Map network drive via script (or migrate share to SharePoint/Azure Files)', support: 'ScriptRequired', confidence: 80, notes: 'No native CSP for drive maps. Use a PowerShell script or, preferably, migrate the share to SharePoint/Azure Files.' },
  { id: 'gpo-printer', match: ['printerconnection', 'deployed printers', 'printer', 'sharedprinter'], gpoName: 'Deployed Printers', scope: 'Both', category: 'Printers', target: 'Manual', intunePath: 'Universal Print + Intune printer provisioning policy', setting: 'Universal Print shared printer assignment', support: 'Manual', confidence: 75, notes: 'Migrate to Universal Print; assign via Entra groups. Direct-IP printers can be scripted.' },
  { id: 'gpo-logon-script', match: ['logon script', 'logonscript', 'scripts (logon/logoff)'], gpoName: 'Logon/Logoff Scripts', scope: 'User', category: 'Scripts', target: 'PowerShell Script', intunePath: 'Devices > Scripts and remediations', setting: 'Platform/PowerShell script', support: 'ScriptRequired', confidence: 78, notes: 'Logon-script logic must be rewritten — Intune scripts run in different context/timing (system vs user, run-once).' },
  { id: 'gpo-mapped-printer-pref', match: ['printers (group policy preferences)', 'printerset'], gpoName: 'Printers (Group Policy Preferences)', scope: 'User', category: 'Preferences - Printers', target: 'Manual', intunePath: 'Universal Print', setting: 'Universal Print', support: 'Manual', confidence: 72 },
  { id: 'gpo-onedrive-kfm', match: ['kfmoptinwithwizard', 'kfmsilentoptin', 'onedrive'], gpoName: 'OneDrive Known Folder Move', scope: 'Computer', category: 'OneDrive', target: 'OneDrive Policy', intunePath: 'Settings Catalog > OneDrive', setting: 'Silently move known folders to OneDrive', support: 'Supported', confidence: 95 },
  { id: 'gpo-edge-homepage', match: ['homepagelocation', 'edge', 'browser home page'], gpoName: 'Browser home page / startup', scope: 'Both', category: 'Microsoft Edge', target: 'Administrative Templates', intunePath: 'Settings Catalog > Microsoft Edge', setting: 'Configure the home page URL', support: 'Supported', confidence: 90 },
  { id: 'gpo-wsus', match: ['wuserver', 'windows update', 'wsus', 'specify intranet'], gpoName: 'Specify intranet Microsoft update service location (WSUS)', scope: 'Computer', category: 'Windows Update', target: 'Settings Catalog', intunePath: 'Devices > Update rings / Settings Catalog > Windows Update', setting: 'Windows Update for Business (replaces WSUS)', support: 'ScriptRequired', confidence: 70, notes: 'WSUS does not map directly — move to Windows Update for Business update rings.' },
  { id: 'gpo-applocker', match: ['applocker', 'srp', 'software restriction'], gpoName: 'AppLocker / Software Restriction Policies', scope: 'Computer', category: 'Application Control', target: 'Endpoint Security', intunePath: 'Endpoint security > Attack surface reduction / App Control for Business', setting: 'App Control / AppLocker CSP', support: 'Supported', confidence: 82 },
  { id: 'gpo-mapped-home', match: ['home directory', 'homedir', 'profile path', 'roaming profile'], gpoName: 'Roaming profiles / Home directory', scope: 'User', category: 'User Profiles', target: 'Unsupported', intunePath: 'Enterprise State Roaming (Entra) + OneDrive KFM', setting: 'No direct equivalent', support: 'Unsupported', confidence: 60, notes: 'Roaming profiles are not supported in cloud-only. Use Enterprise State Roaming + OneDrive KFM; expect redesign.' },
  { id: 'gpo-ie', match: ['internet explorer', 'iexplore', 'ie maintenance'], gpoName: 'Internet Explorer Maintenance', scope: 'Both', category: 'Internet Explorer', target: 'Unsupported', intunePath: 'Edge IE Mode site list', setting: 'Internet Explorer is retired', support: 'Unsupported', confidence: 65, notes: 'IE is end-of-life. Map required sites to Edge IE Mode; discard the rest.' },
  { id: 'gpo-software-install', match: ['software installation', 'msi deploy', 'assigned applications'], gpoName: 'Software Installation (assigned/published MSI)', scope: 'Both', category: 'Software Installation', target: 'Win32 App', intunePath: 'Apps > Windows > Win32 app (.intunewin)', setting: 'Repackage MSI as Win32/LOB app', support: 'ScriptRequired', confidence: 80, notes: 'Repackage as Win32 (.intunewin) or LOB MSI; assign to Entra groups.' },
];

export const gpoTargets: GpoTarget[] = ['Settings Catalog', 'Administrative Templates', 'Security Baseline', 'Endpoint Security', 'PowerShell Script', 'Win32 App', 'OneDrive Policy', 'Defender Policy', 'Manual', 'Unsupported'];

export const supportColor: Record<GpoSupport, string> = {
  Supported: 'green', ScriptRequired: 'orange', Manual: 'purple', Unsupported: 'red',
};
