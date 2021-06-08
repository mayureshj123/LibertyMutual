

import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface ITaskListWebpartProps {
  description: string;
  context : WebPartContext;
  // lists: string | string[]; // Stores the list ID(s)

}
