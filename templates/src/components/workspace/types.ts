import type { AttachmentPayload } from '../../lib/api';

export type ComposeMode = 'new' | 'reply' | 'replyAll' | 'forward';

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
