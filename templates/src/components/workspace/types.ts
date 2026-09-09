import { format } from 'date-fns';
import type { AttachmentPayload } from '../../lib/api';
import type { Email } from '../../store/useAppStore';

export type ComposeMode = 'new' | 'reply' | 'replyAll' | 'forward';
export type ViewMode = 'html' | 'text' | 'headers';

export interface MetaFormState {
  tags: string;
  followUp: string;
  notes: string;
  snoozedUntil: string;
  status: string;
}

export interface ComposeFormState {
  open: boolean;
  mode: ComposeMode;
  messageId: string;
  draftMessageId: string;
  subject: string;
  to: string;
  cc: string;
  bcc: string;
  bodyText: string;
  attachments: AttachmentPayload[];
  attachmentNames: string[];
  submitting: boolean;
}

export interface RuleEditorState {
  id: number | null;
  name: string;
  enabled: boolean;
  priority: number;
  conditionsText: string;
  actionsText: string;
}

export const EMPTY_COMPOSE: ComposeFormState = {
  open: false,
  mode: 'new',
  messageId: '',
  draftMessageId: '',
  subject: '',
  to: '',
  cc: '',
  bcc: '',
  bodyText: '',
  attachments: [],
  attachmentNames: [],
  submitting: false,
};

export const EMPTY_RULE_EDITOR: RuleEditorState = {
  id: null,
  name: '',
  enabled: true,
  priority: 100,
  conditionsText: '{\n  "subject_contains": ""\n}',
  actionsText: '{\n  "mark_read": true,\n  "tags": [""]\n}',
};

export const EMPTY_META_FORM: MetaFormState = {
  tags: '',
  followUp: '',
  notes: '',
  snoozedUntil: '',
  status: 'active',
};

export function formatMailDate(value: string | undefined | null, pattern = 'yyyy-MM-dd HH:mm'): string {
  if (!value) {
    return '-';
  }
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) {
    return '-';
  }
  try {
    return format(date, pattern);
  } catch {
    return '-';
  }
}

export function isDraftFolder(email: Email | null, folderId: string) {
  if (!email) {
    return false;
  }
  const normalized = (email.folderId || folderId || '').toLowerCase();
  return normalized.includes('draft');
}
