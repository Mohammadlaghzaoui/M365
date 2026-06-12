// Catalog of common GPO settings you can pick to BUILD a custom GPO.
// Each maps to a registry value the generated PowerShell will set.

export interface GpoSettingDef {
  id: string;
  category: string;
  label: string;
  hive: 'HKLM' | 'HKCU';
  key: string;
  valueName: string;
  type: 'DWord' | 'String' | 'ExpandString';
  /** input kind for the UI */
  input: 'toggle' | 'number' | 'text' | 'choice';
  choices?: { label: string; value: string | number }[];
  /** value used when a toggle is ON (enabled) */
  onValue?: string | number;
  offValue?: string | number;
  default?: string | number;
  help?: string;
}

export const gpoCatalog: GpoSettingDef[] = [
  // --- Security / lockdown ---
  { id: 'disable-cmd', category: 'Security', label: 'Prevent access to Command Prompt', hive: 'HKCU', key: 'Software\\Policies\\Microsoft\\Windows\\System', valueName: 'DisableCMD', type: 'DWord', input: 'toggle', onValue: 2, offValue: 0, help: '2 = block CMD and scripts' },
  { id: 'disable-regedit', category: 'Security', label: 'Prevent access to Registry editing tools', hive: 'HKCU', key: 'Software\\Microsoft\\Windows\\CurrentVersion\\Policies\\System', valueName: 'DisableRegistryTools', type: 'DWord', input: 'toggle', onValue: 1, offValue: 0 },
  { id: 'no-control-panel', category: 'Security', label: 'Prohibit access to Control Panel', hive: 'HKCU', key: 'Software\\Microsoft\\Windows\\CurrentVersion\\Policies\\Explorer', valueName: 'NoControlPanel', type: 'DWord', input: 'toggle', onValue: 1, offValue: 0 },
  { id: 'no-run', category: 'Security', label: 'Remove Run menu from Start Menu', hive: 'HKCU', key: 'Software\\Microsoft\\Windows\\CurrentVersion\\Policies\\Explorer', valueName: 'NoRun', type: 'DWord', input: 'toggle', onValue: 1, offValue: 0 },
  { id: 'no-task-mgr', category: 'Security', label: 'Disable Task Manager', hive: 'HKCU', key: 'Software\\Microsoft\\Windows\\CurrentVersion\\Policies\\System', valueName: 'DisableTaskMgr', type: 'DWord', input: 'toggle', onValue: 1, offValue: 0 },
  { id: 'no-usb-storage', category: 'Security', label: 'Disable USB storage devices', hive: 'HKLM', key: 'SYSTEM\\CurrentControlSet\\Services\\USBSTOR', valueName: 'Start', type: 'DWord', input: 'toggle', onValue: 4, offValue: 3, help: '4 = disabled, 3 = enabled' },
  { id: 'disable-autorun', category: 'Security', label: 'Disable AutoRun on all drives', hive: 'HKLM', key: 'Software\\Microsoft\\Windows\\CurrentVersion\\Policies\\Explorer', valueName: 'NoDriveTypeAutoRun', type: 'DWord', input: 'toggle', onValue: 255, offValue: 0 },
  { id: 'require-ctrlaltdel', category: 'Security', label: 'Require CTRL+ALT+DEL at logon', hive: 'HKLM', key: 'Software\\Microsoft\\Windows\\CurrentVersion\\Policies\\System', valueName: 'DisableCAD', type: 'DWord', input: 'toggle', onValue: 0, offValue: 1 },

  // --- Screen / session ---
  { id: 'screensaver-timeout', category: 'Screen & session', label: 'Screen saver timeout (seconds)', hive: 'HKCU', key: 'Software\\Policies\\Microsoft\\Windows\\Control Panel\\Desktop', valueName: 'ScreenSaveTimeOut', type: 'String', input: 'number', default: 900 },
  { id: 'screensaver-secure', category: 'Screen & session', label: 'Password-protect the screen saver', hive: 'HKCU', key: 'Software\\Policies\\Microsoft\\Windows\\Control Panel\\Desktop', valueName: 'ScreenSaverIsSecure', type: 'String', input: 'toggle', onValue: '1', offValue: '0' },
  { id: 'inactivity-lock', category: 'Screen & session', label: 'Machine inactivity limit before lock (seconds)', hive: 'HKLM', key: 'Software\\Microsoft\\Windows\\CurrentVersion\\Policies\\System', valueName: 'InactivityTimeoutSecs', type: 'DWord', input: 'number', default: 900 },
  { id: 'legal-notice-caption', category: 'Screen & session', label: 'Logon banner title', hive: 'HKLM', key: 'Software\\Microsoft\\Windows\\CurrentVersion\\Policies\\System', valueName: 'legalnoticecaption', type: 'String', input: 'text', default: 'Authorized use only' },
  { id: 'legal-notice-text', category: 'Screen & session', label: 'Logon banner message', hive: 'HKLM', key: 'Software\\Microsoft\\Windows\\CurrentVersion\\Policies\\System', valueName: 'legalnoticetext', type: 'String', input: 'text', default: 'This system is for authorized users only.' },

  // --- Windows Update ---
  { id: 'wsus-server', category: 'Windows Update', label: 'WSUS server URL', hive: 'HKLM', key: 'Software\\Policies\\Microsoft\\Windows\\WindowsUpdate', valueName: 'WUServer', type: 'String', input: 'text', default: 'http://wsus.contoso.local:8530' },
  { id: 'au-option', category: 'Windows Update', label: 'Automatic Updates behaviour', hive: 'HKLM', key: 'Software\\Policies\\Microsoft\\Windows\\WindowsUpdate\\AU', valueName: 'AUOptions', type: 'DWord', input: 'choice', choices: [{ label: 'Notify download', value: 2 }, { label: 'Auto download, notify install', value: 3 }, { label: 'Auto download + schedule install', value: 4 }], default: 4 },
  { id: 'no-auto-reboot', category: 'Windows Update', label: 'No auto-restart with logged-on users', hive: 'HKLM', key: 'Software\\Policies\\Microsoft\\Windows\\WindowsUpdate\\AU', valueName: 'NoAutoRebootWithLoggedOnUsers', type: 'DWord', input: 'toggle', onValue: 1, offValue: 0 },

  // --- Microsoft Defender ---
  { id: 'defender-rtp', category: 'Defender', label: 'Real-time protection', hive: 'HKLM', key: 'Software\\Policies\\Microsoft\\Windows Defender\\Real-Time Protection', valueName: 'DisableRealtimeMonitoring', type: 'DWord', input: 'toggle', onValue: 0, offValue: 1, help: 'ON = protection enabled (DisableRealtimeMonitoring=0)' },
  { id: 'defender-cloud', category: 'Defender', label: 'Cloud-delivered protection (MAPS)', hive: 'HKLM', key: 'Software\\Policies\\Microsoft\\Windows Defender\\Spynet', valueName: 'SpynetReporting', type: 'DWord', input: 'choice', choices: [{ label: 'Off', value: 0 }, { label: 'Basic', value: 1 }, { label: 'Advanced', value: 2 }], default: 2 },

  // --- Explorer / OneDrive ---
  { id: 'kfm-onedrive', category: 'OneDrive', label: 'Silently move known folders to OneDrive (tenant ID)', hive: 'HKLM', key: 'Software\\Policies\\Microsoft\\OneDrive', valueName: 'KFMSilentOptIn', type: 'String', input: 'text', default: '<your-tenant-guid>' },
  { id: 'hide-recently-added', category: 'Explorer', label: 'Hide recently added apps in Start', hive: 'HKLM', key: 'Software\\Policies\\Microsoft\\Windows\\Explorer', valueName: 'HideRecentlyAddedApps', type: 'DWord', input: 'toggle', onValue: 1, offValue: 0 },

  // --- Edge ---
  { id: 'edge-homepage', category: 'Microsoft Edge', label: 'Edge home page URL', hive: 'HKLM', key: 'Software\\Policies\\Microsoft\\Edge', valueName: 'HomepageLocation', type: 'String', input: 'text', default: 'https://intranet.contoso.com' },
  { id: 'edge-startup', category: 'Microsoft Edge', label: 'Edge startup behaviour', hive: 'HKLM', key: 'Software\\Policies\\Microsoft\\Edge', valueName: 'RestoreOnStartup', type: 'DWord', input: 'choice', choices: [{ label: 'New tab page', value: 5 }, { label: 'Restore last session', value: 1 }, { label: 'Open specific pages', value: 4 }], default: 4 },
];

export const gpoCategories = [...new Set(gpoCatalog.map((c) => c.category))];
