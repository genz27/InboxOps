import { EmailAccount, Folder, Email } from '../store/useAppStore';

export const mockAccounts: EmailAccount[] = [
  {
    id: '1',
    label: 'Main Admin',
    email: 'admin@example.com',
    method: 'graph_api',
    notes: 'Primary account',
    status: 'connected',
  },
  {
    id: '2',
    label: 'Support',
    email: 'support@example.com',
    method: 'imap_new',
    notes: 'Support mailbox',
    status: 'connected',
  },
  {
    id: '3',
    label: 'Sales',
    email: 'sales@example.com',
    method: 'imap_old',
    notes: 'Sales team mailbox',
    status: 'disconnected',
  },
];

export const mockFolders: Folder[] = [
  { id: 'f1', name: 'Inbox', displayName: 'Inbox', unreadCount: 2, totalCount: 15, type: 'inbox' },
  { id: 'f2', name: 'Sent', displayName: 'Sent', unreadCount: 0, totalCount: 45, type: 'sent' },
  { id: 'f3', name: 'Drafts', displayName: 'Drafts', unreadCount: 0, totalCount: 3, type: 'drafts' },
  { id: 'f4', name: 'Spam', displayName: 'Spam', unreadCount: 5, totalCount: 12, type: 'spam' },
  { id: 'f5', name: 'Trash', displayName: 'Trash', unreadCount: 0, totalCount: 8, type: 'trash' },
  { id: 'f6', name: 'Archive', displayName: 'Archive', unreadCount: 0, totalCount: 120, type: 'archive' },
];

export const mockEmails: Email[] = Array.from({ length: 50 }).map((_, i) => ({
  id: `e${i}`,
  subject: `Test Email Subject ${i + 1}`,
  sender: `sender${i}@example.com`,
  sender_name: `Sender ${i + 1}`,
  to_recipients: ['admin@example.com'],
  cc_recipients: [],
  bcc_recipients: [],
  date: new Date(Date.now() - i * 3600000 * 24).toISOString(),
  is_read: i % 3 !== 0,
  is_flagged: i % 5 === 0,
  importance: i % 7 === 0 ? 'high' : 'normal',
  has_attachments: i % 4 === 0,
  preview: `This is a preview of email ${i + 1}...`,
  body_text: `Hello,\n\nThis is the plain text body of email ${i + 1}.\n\nBest,\nSender`,
  body_html: `<div><p>Hello,</p><p>This is the <b>HTML</b> body of email ${i + 1}.</p><p>Best,<br>Sender</p></div>`,
  headers: `Received: by mx.google.com\nDate: ${new Date().toUTCString()}\nFrom: sender${i}@example.com\nTo: admin@example.com`,
  internet_message_id: `<msg-${i}@example.com>`,
  conversation_id: `conv-${Math.floor(i / 3)}`,
  attachments:
    i % 4 === 0
      ? [
          {
            id: `att-${i}`,
            name: `document_${i}.pdf`,
            contentType: 'application/pdf',
            size: 1024 * 1024 * (i + 1),
            isInline: false,
          },
        ]
      : [],
  folderId: i < 15 ? 'f1' : i < 30 ? 'f2' : i < 40 ? 'f6' : 'f5',
  method: i % 2 === 0 ? 'graph_api' : 'imap_new',
}));
