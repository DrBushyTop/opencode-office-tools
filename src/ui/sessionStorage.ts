import type { ModelType } from "./components/HeaderBar";
import type { Message } from "./components/MessageList";

export interface SavedSession {
  id: string;
  title: string;
  model: ModelType;
  messages: Message[];
  createdAt: string;
  updatedAt: string;
}

export type OfficeHost = "powerpoint" | "word" | "excel";

export function getHostFromOfficeHost(host: typeof Office.HostType[keyof typeof Office.HostType]): OfficeHost {
  switch (host) {
    case Office.HostType.PowerPoint:
      return "powerpoint";
    case Office.HostType.Word:
      return "word";
    case Office.HostType.Excel:
      return "excel";
    default:
      return "word";
  }
}
