import { KeyRound, Mail, Boxes, Users } from 'lucide-react';
import { AssistantPage } from '../components/AssistantPage';
import { entraWorkflows } from '../data/entraWorkflows';
import { exchangeWorkflows } from '../data/exchangeWorkflows';
import { sharepointWorkflows } from '../data/sharepointWorkflows';
import { teamsWorkflows } from '../data/teamsWorkflows';

export function EntraAssistant() {
  return <AssistantPage title="Entra ID Assistant" subtitle="Guided workflows for identity, sign-in, MFA, Conditional Access and account issues." icon={<KeyRound size={20} />} workflows={entraWorkflows} />;
}

export function ExchangeAssistant() {
  return <AssistantPage title="Exchange Online Assistant" subtitle="Guided workflows for mail flow, mailboxes, permissions, rooms and email authentication." icon={<Mail size={20} />} workflows={exchangeWorkflows} />;
}

export function SharePointAssistant() {
  return <AssistantPage title="SharePoint Online Assistant" subtitle="Guided workflows for access, sharing, restores, storage and OneDrive sync." icon={<Boxes size={20} />} workflows={sharepointWorkflows} />;
}

export function TeamsAssistant() {
  return <AssistantPage title="Teams Assistant" subtitle="Guided workflows for Teams access, meetings, channels, guests, devices and policies." icon={<Users size={20} />} workflows={teamsWorkflows} />;
}
