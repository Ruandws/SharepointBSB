import { IMenuItem } from '../DashboardWebPart';

export interface IDashboardProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  menuItems: IMenuItem[];
}