import type { Message } from "./components/MessageList";
import type { ModelType } from "./components/HeaderBar";

export interface SavedSession {
  id: string;
  title: string;
  model: ModelType;
  messages: Message[];
  createdAt: string;
  updatedAt: string;
}

export type OfficeHost = "powerpoint" | "word" | "excel";

const MAX_SESSIONS_PER_HOST = 50;

function getStorageKey(host: OfficeHost): string {
  return `opencode-sessions-${host}`;
}

export function getSavedSessions(host: OfficeHost): SavedSession[] {
  try {
    const stored = localStorage.getItem(getStorageKey(host));
    return stored ? JSON.parse(stored) : [];
  } catch {
    return [];
  }
}

export function saveSession(
  host: OfficeHost,
  session: SavedSession
): void {
  try {
    const sessions = getSavedSessions(host);
    
    // Check if session already exists (update it)
    const existingIndex = sessions.findIndex(s => s.id === session.id);
    if (existingIndex >= 0) {
      sessions[existingIndex] = session;
    } else {
      // Add new session at the beginning
      sessions.unshift(session);
    }
    
    // Keep only the last MAX_SESSIONS_PER_HOST
    const trimmed = sessions.slice(0, MAX_SESSIONS_PER_HOST);
    
    localStorage.setItem(getStorageKey(host), JSON.stringify(trimmed));
  } catch (e) {
    console.error("Failed to save session:", e);
  }
}

export function deleteSession(host: OfficeHost, sessionId: string): void {
  try {
    const sessions = getSavedSessions(host);
    const filtered = sessions.filter(s => s.id !== sessionId);
    localStorage.setItem(getStorageKey(host), JSON.stringify(filtered));
  } catch (e) {
    console.error("Failed to delete session:", e);
  }
}

export function generateSessionTitle(messages: Message[]): string {
  // Find first user message to use as title
  const firstUserMessage = messages.find(m => m.sender === "user");
  if (firstUserMessage?.text) {
    // Truncate to reasonable length
    const text = firstUserMessage.text.trim();
    if (text.length <= 50) return text;
    return text.substring(0, 47) + "...";
  }
  return "New conversation";
}

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
