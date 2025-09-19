#!/usr/bin/env node

import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import {
  CallToolRequestSchema,
  ErrorCode,
  ListToolsRequestSchema,
  McpError,
} from '@modelcontextprotocol/sdk/types.js';
import fs from 'fs/promises';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
import os from 'os';
import http from 'http';
import { google } from 'googleapis';

const SCOPES = [
  'https://www.googleapis.com/auth/gmail.modify',
  'https://www.googleapis.com/auth/gmail.send',
  'https://www.googleapis.com/auth/gmail.labels'
];

function base64UrlEncode(str) {
  return Buffer.from(str)
    .toString('base64')
    .replace(/\+/g, '-')
    .replace(/\//g, '_')
    .replace(/=+$/, '');
}

function decodeBase64Url(data) {
  const buff = Buffer.from(data.replace(/-/g, '+').replace(/_/g, '/'), 'base64');
  return buff.toString('utf-8');
}

class GmailMcpServer {
  constructor() {
    this.server = new Server({
      name: 'gmail-mcp',
      version: '1.0.0',
    }, {
      capabilities: {
        tools: {},
      },
    });

    this.oauthState = {
      server: null,
      pending: false,
      redirectPort: 53682,
    };

    this.uiState = {
      server: null,
      port: 53750,
    };

    this.setupToolHandlers();
  }

  formatContent(content) {
    if (Array.isArray(content)) {
      return content.map(item => ({ type: 'text', text: typeof item === 'string' ? item : JSON.stringify(item, null, 2) }));
    }
    return [{ type: 'text', text: typeof content === 'string' ? content : JSON.stringify(content, null, 2) }];
  }

  getConfigDir() {
    // Use project directory instead of global config
    return path.join(__dirname);
  }

  async ensureConfigDir() {
    const dir = this.getConfigDir();
    try {
      await fs.mkdir(dir, { recursive: true });
    } catch {}
    return dir;
  }

  getRulesConfigPath() {
    return path.join(this.getConfigDir(), 'auto-labeling-rules.json');
  }

  getIgnoredLabelsConfigPath() {
    return path.join(this.getConfigDir(), 'ignored-labels-inbox.json');
  }

  async loadAutoLabelingRules() {
    try {
      const rulesPath = this.getRulesConfigPath();
      const raw = await fs.readFile(rulesPath, 'utf8');
      const config = JSON.parse(raw);
      return config.rules || [];
    } catch {
      // Return empty array if file doesn't exist or is invalid
      return [];
    }
  }

  async loadIgnoredLabels() {
    try {
      const ignoredLabelsPath = this.getIgnoredLabelsConfigPath();
      const raw = await fs.readFile(ignoredLabelsPath, 'utf8');
      const config = JSON.parse(raw);
      return config.labels || [];
    } catch {
      // Return empty array if file doesn't exist or is invalid
      return [];
    }
  }

  async saveAutoLabelingRules(rules) {
    const rulesPath = this.getRulesConfigPath();
    const config = {
      version: '1.0.0',
      lastUpdated: new Date().toISOString(),
      rules: rules
    };
    await fs.writeFile(rulesPath, JSON.stringify(config, null, 2), 'utf8');
  }

  async loadCredentials() {
    // Prefer environment variables
    const clientId = process.env.GOOGLE_CLIENT_ID;
    const clientSecret = process.env.GOOGLE_CLIENT_SECRET;
    if (clientId && clientSecret) {
      return { client_id: clientId, client_secret: clientSecret };
    }

    // Fallback to credentials.json in config dir
    const dir = await this.ensureConfigDir();
    const credPath = path.join(dir, 'credentials.json');
    try {
      const raw = await fs.readFile(credPath, 'utf8');
      const json = JSON.parse(raw);
      if (json.installed) {
        return {
          client_id: json.installed.client_id,
          client_secret: json.installed.client_secret,
        };
      }
      return {
        client_id: json.client_id,
        client_secret: json.client_secret,
      };
    } catch (e) {
      throw new McpError(ErrorCode.InvalidRequest, 'Missing Google OAuth credentials. Set GOOGLE_CLIENT_ID and GOOGLE_CLIENT_SECRET env vars or place credentials.json in ~/.config/mcp-gmail');
    }
  }

  async getOAuth2Client() {
    const { client_id, client_secret } = await this.loadCredentials();
    const redirectUri = `http://127.0.0.1:${this.oauthState.redirectPort}/oauth2callback`;
    const oAuth2Client = new google.auth.OAuth2(client_id, client_secret, redirectUri);

    const token = await this.loadToken();
    if (token) {
      oAuth2Client.setCredentials(token);
    }
    return oAuth2Client;
  }

  async loadToken() {
    try {
      const tokenPath = path.join(await this.ensureConfigDir(), 'token.json');
      const raw = await fs.readFile(tokenPath, 'utf8');
      return JSON.parse(raw);
    } catch {
      return null;
    }
  }

  async saveToken(token) {
    const tokenPath = path.join(await this.ensureConfigDir(), 'token.json');
    await fs.writeFile(tokenPath, JSON.stringify(token, null, 2), 'utf8');
  }

  async getGmail() {
    const auth = await this.getOAuth2Client();
    const token = await this.loadToken();
    if (!token) {
      throw new McpError(ErrorCode.InvalidRequest, 'Not authorized. Run start_oauth to authorize this server.');
    }
    return google.gmail({ version: 'v1', auth });
  }

  async startLocalOAuthServer(oAuth2Client, port) {
    if (this.oauthState.pending) {
      throw new McpError(ErrorCode.InvalidRequest, 'OAuth flow already in progress. Complete it in your browser.');
    }

    this.oauthState.pending = true;

    const server = http.createServer(async (req, res) => {
      try {
        const url = new URL(req.url, `http://127.0.0.1:${port}`);
        if (url.pathname !== '/oauth2callback') {
          res.writeHead(404);
          res.end('Not Found');
          return;
        }
        const code = url.searchParams.get('code');
        if (!code) {
          res.writeHead(400);
          res.end('Missing code');
          return;
        }
        const { tokens } = await oAuth2Client.getToken(code);
        await this.saveToken(tokens);
        oAuth2Client.setCredentials(tokens);
        res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' });
        res.end('<p>✅ Authorization complete. You can close this window.</p>');
      } catch (e) {
        res.writeHead(500);
        res.end('Auth error');
      } finally {
        this.oauthState.pending = false;
        setTimeout(() => {
          try { server.close(); } catch {}
          this.oauthState.server = null;
        }, 100);
      }
    });

    await new Promise((resolve, reject) => {
      server.once('error', reject);
      server.listen(port, '127.0.0.1', () => resolve());
    });

    this.oauthState.server = server;
  }

  setupToolHandlers() {
    this.server.setRequestHandler(ListToolsRequestSchema, async () => ({
      tools: [
        {
          name: 'start_oauth',
          description: 'Start OAuth flow. Returns a URL to open in your browser to authorize Gmail access.',
          inputSchema: {
            type: 'object',
            properties: {
              port: { type: 'number', description: 'Local port to listen on (optional)', default: 53682 },
            },
          },
        },
        {
          name: 'auth_status',
          description: 'Check whether the server is authorized with Google and which email is active.',
          inputSchema: { type: 'object', properties: {} },
        },
        {
          name: 'list_unread',
          description: 'List recent unread emails (From, Subject, Date, Snippet).',
          inputSchema: {
            type: 'object',
            properties: {
              max: { type: 'number', description: 'Max emails to list (default 10)', default: 10 },
            },
          },
        },
        {
          name: 'list_recent_unread',
          description: 'List recent unread emails in table format. Defaults to 7 days, accepts days parameter (e.g., "2 days", "3d", "--2").',
          inputSchema: {
            type: 'object',
            properties: {
              days: { type: 'string', description: 'Number of days to look back (e.g., "2", "3d", "2 days", "--2"). Default is "7".', default: '7' },
              max: { type: 'number', description: 'Max emails to list (default 50)', default: 50 },
            },
          },
        },
        {
          name: 'get_message',
          description: 'Get a message by ID with headers, snippet, and plain text body.',
          inputSchema: {
            type: 'object',
            properties: {
              id: { type: 'string', description: 'Gmail message ID' },
            },
            required: ['id'],
          },
        },
        {
          name: 'reply_to_message',
          description: 'Reply to a message by ID. Uses the original sender as the recipient.',
          inputSchema: {
            type: 'object',
            properties: {
              message_id: { type: 'string', description: 'Gmail message ID to reply to' },
              body: { type: 'string', description: 'Plain text reply body' },
            },
            required: ['message_id', 'body'],
          },
        },
        {
          name: 'mark_as_read',
          description: 'Mark a message as read (remove UNREAD label).',
          inputSchema: {
            type: 'object',
            properties: {
              id: { type: 'string', description: 'Gmail message ID' },
            },
            required: ['id'],
          },
        },
        {
          name: 'batch_archive',
          description: 'Archive multiple messages by removing INBOX label.',
          inputSchema: {
            type: 'object',
            properties: {
              ids: { type: 'array', items: { type: 'string' }, description: 'Array of Gmail message IDs' },
            },
            required: ['ids'],
          },
        },
        {
          name: 'batch_delete',
          description: 'Move multiple messages to Trash (safer than permanent delete).',
          inputSchema: {
            type: 'object',
            properties: {
              ids: { type: 'array', items: { type: 'string' }, description: 'Array of Gmail message IDs' },
              permanent: { type: 'boolean', description: 'Permanently delete instead of trash', default: false },
            },
            required: ['ids'],
          },
        },
        {
          name: 'start_ui',
          description: 'Start a minimal local UI to browse unread and take actions.',
          inputSchema: {
            type: 'object',
            properties: {
              port: { type: 'number', description: 'Local port for the UI server', default: 53750 },
              query: { type: 'string', description: 'Gmail search query for listing', default: 'is:unread' },
              max: { type: 'number', description: 'Max messages to display', default: 50 }
            },
          },
        },
        {
          name: 'stop_ui',
          description: 'Stop the currently running UI server.',
          inputSchema: {
            type: 'object',
            properties: {},
          },
        },
        {
          name: 'search_emails',
          description: 'Search emails by subject, sender name, or email address. Supports Gmail search syntax.',
          inputSchema: {
            type: 'object',
            properties: {
              query: { type: 'string', description: 'Search query (e.g., "capterra", "from:example.com", "subject:urgent")' },
              max: { type: 'number', description: 'Max emails to return (default 10)', default: 10 },
            },
            required: ['query'],
          },
        },
        {
          name: 'delete_email',
          description: 'Delete a single email by ID. Moves to trash by default.',
          inputSchema: {
            type: 'object',
            properties: {
              id: { type: 'string', description: 'Gmail message ID' },
              permanent: { type: 'boolean', description: 'Permanently delete instead of trash', default: false },
            },
            required: ['id'],
          },
        },
        {
          name: 'delete_emails_by_query',
          description: 'Delete multiple emails matching a search query.',
          inputSchema: {
            type: 'object',
            properties: {
              query: { type: 'string', description: 'Search query to find emails to delete' },
              max: { type: 'number', description: 'Max emails to delete (default 50)', default: 50 },
              permanent: { type: 'boolean', description: 'Permanently delete instead of trash', default: false },
              dry_run: { type: 'boolean', description: 'Show what would be deleted without actually deleting', default: false },
            },
            required: ['query'],
          },
        },
        {
          name: 'bulk_delete_emails',
          description: 'High-performance bulk delete for large numbers of emails with progress reporting.',
          inputSchema: {
            type: 'object',
            properties: {
              query: { type: 'string', description: 'Search query to find emails to delete' },
              batch_size: { type: 'number', description: 'Batch size for processing (default 100)', default: 100 },
              permanent: { type: 'boolean', description: 'Permanently delete instead of trash', default: false },
              dry_run: { type: 'boolean', description: 'Show what would be deleted without actually deleting', default: false },
            },
            required: ['query'],
          },
        },
        {
          name: 'create_draft',
          description: 'Create a draft email without sending it. Perfect for collaborative email crafting.',
          inputSchema: {
            type: 'object',
            properties: {
              to: { type: 'string', description: 'Recipient email address' },
              subject: { type: 'string', description: 'Email subject' },
              body: { type: 'string', description: 'Plain text email body' },
              html: { type: 'string', description: 'HTML email body (optional)' },
              reply_to_message_id: { type: 'string', description: 'Message ID to reply to (optional)' },
            },
            required: ['to', 'subject', 'body'],
          },
        },
        {
          name: 'create_reply_draft',
          description: 'Create a proper reply draft with threading. Reply to a message by ID and create a draft without sending.',
          inputSchema: {
            type: 'object',
            properties: {
              message_id: { type: 'string', description: 'Gmail message ID to reply to' },
              body: { type: 'string', description: 'Plain text reply body' },
              html: { type: 'string', description: 'HTML reply body (optional)' },
            },
            required: ['message_id', 'body'],
          },
        },
        {
          name: 'get_draft',
          description: 'Get a draft email by ID to review its content before sending.',
          inputSchema: {
            type: 'object',
            properties: {
              draft_id: { type: 'string', description: 'Draft email ID' },
            },
            required: ['draft_id'],
          },
        },
        {
          name: 'list_drafts',
          description: 'List all draft emails for review and management.',
          inputSchema: {
            type: 'object',
            properties: {
              max: { type: 'number', description: 'Maximum number of drafts to return (default 10)', default: 10 },
            },
          },
        },
        {
          name: 'update_draft',
          description: 'Update an existing draft email with new content.',
          inputSchema: {
            type: 'object',
            properties: {
              draft_id: { type: 'string', description: 'Draft email ID to update' },
              to: { type: 'string', description: 'Recipient email address' },
              subject: { type: 'string', description: 'Email subject' },
              body: { type: 'string', description: 'Plain text email body' },
              html: { type: 'string', description: 'HTML email body (optional)' },
            },
            required: ['draft_id'],
          },
        },
        {
          name: 'send_draft',
          description: 'Send a draft email after reviewing and approving its content.',
          inputSchema: {
            type: 'object',
            properties: {
              draft_id: { type: 'string', description: 'Draft email ID to send' },
            },
            required: ['draft_id'],
          },
        },
        {
          name: 'delete_draft',
          description: 'Delete a draft email without sending it.',
          inputSchema: {
            type: 'object',
            properties: {
              draft_id: { type: 'string', description: 'Draft email ID to delete' },
            },
            required: ['draft_id'],
          },
        },
        {
          name: 'snooze_email',
          description: 'Snooze an email to reappear in inbox at a specified date and time.',
          inputSchema: {
            type: 'object',
            properties: {
              id: { type: 'string', description: 'Gmail message ID to snooze' },
              snooze_date: { type: 'string', description: 'Date and time to snooze until (ISO format: YYYY-MM-DDTHH:MM:SS or relative: "tomorrow 9am", "next monday", "in 2 hours")' },
            },
            required: ['id', 'snooze_date'],
          },
        },
        {
          name: 'unsnooze_email',
          description: 'Unsnooze an email by moving it back to inbox and removing snooze label.',
          inputSchema: {
            type: 'object',
            properties: {
              id: { type: 'string', description: 'Gmail message ID to unsnooze' },
            },
            required: ['id'],
          },
        },
        {
          name: 'list_snoozed_emails',
          description: 'List all snoozed emails (emails with SNOOZED label).',
          inputSchema: {
            type: 'object',
            properties: {
              max: { type: 'number', description: 'Max emails to list (default 20)', default: 20 },
            },
          },
        },
        // Label Management Tools
        {
          name: 'create_label',
          description: 'Create a single Gmail label.',
          inputSchema: {
            type: 'object',
            properties: {
              name: { type: 'string', description: 'Label name' },
              label_list_visibility: { type: 'string', description: 'Label visibility in label list (labelShow, labelHide)', default: 'labelShow' },
              message_list_visibility: { type: 'string', description: 'Message visibility in message list (show, hide)', default: 'show' },
            },
            required: ['name'],
          },
        },
        {
          name: 'create_labels',
          description: 'Create multiple Gmail labels at once.',
          inputSchema: {
            type: 'object',
            properties: {
              labels: { 
                type: 'array', 
                items: { 
                  type: 'object',
                  properties: {
                    name: { type: 'string', description: 'Label name' },
                    label_list_visibility: { type: 'string', description: 'Label visibility in label list (labelShow, labelHide)', default: 'labelShow' },
                    message_list_visibility: { type: 'string', description: 'Message visibility in message list (show, hide)', default: 'show' },
                  },
                  required: ['name']
                },
                description: 'Array of label objects to create' 
              },
            },
            required: ['labels'],
          },
        },
        {
          name: 'list_labels',
          description: 'List all Gmail labels.',
          inputSchema: {
            type: 'object',
            properties: {
              include_system: { type: 'boolean', description: 'Include system labels (INBOX, SENT, etc.)', default: true },
            },
          },
        },
        {
          name: 'get_label',
          description: 'Get details of a specific Gmail label.',
          inputSchema: {
            type: 'object',
            properties: {
              label_id: { type: 'string', description: 'Gmail label ID' },
            },
            required: ['label_id'],
          },
        },
        {
          name: 'update_label',
          description: 'Update label properties (name, visibility).',
          inputSchema: {
            type: 'object',
            properties: {
              label_id: { type: 'string', description: 'Gmail label ID' },
              name: { type: 'string', description: 'New label name' },
              label_list_visibility: { type: 'string', description: 'Label visibility in label list (labelShow, labelHide)' },
              message_list_visibility: { type: 'string', description: 'Message visibility in message list (show, hide)' },
            },
            required: ['label_id'],
          },
        },
        {
          name: 'delete_label',
          description: 'Delete a Gmail label.',
          inputSchema: {
            type: 'object',
            properties: {
              label_id: { type: 'string', description: 'Gmail label ID' },
            },
            required: ['label_id'],
          },
        },
        {
          name: 'label_email',
          description: 'Add one or more labels to a single email.',
          inputSchema: {
            type: 'object',
            properties: {
              email_id: { type: 'string', description: 'Gmail message ID' },
              label_ids: { type: 'array', items: { type: 'string' }, description: 'Array of label IDs to add' },
            },
            required: ['email_id', 'label_ids'],
          },
        },
        {
          name: 'label_emails',
          description: 'Add labels to multiple emails.',
          inputSchema: {
            type: 'object',
            properties: {
              email_ids: { type: 'array', items: { type: 'string' }, description: 'Array of Gmail message IDs' },
              label_ids: { type: 'array', items: { type: 'string' }, description: 'Array of label IDs to add' },
            },
            required: ['email_ids', 'label_ids'],
          },
        },
        {
          name: 'unlabel_email',
          description: 'Remove one or more labels from a single email.',
          inputSchema: {
            type: 'object',
            properties: {
              email_id: { type: 'string', description: 'Gmail message ID' },
              label_ids: { type: 'array', items: { type: 'string' }, description: 'Array of label IDs to remove' },
            },
            required: ['email_id', 'label_ids'],
          },
        },
        {
          name: 'unlabel_emails',
          description: 'Remove labels from multiple emails.',
          inputSchema: {
            type: 'object',
            properties: {
              email_ids: { type: 'array', items: { type: 'string' }, description: 'Array of Gmail message IDs' },
              label_ids: { type: 'array', items: { type: 'string' }, description: 'Array of label IDs to remove' },
            },
            required: ['email_ids', 'label_ids'],
          },
        },
        {
          name: 'set_email_labels',
          description: 'Replace all labels on an email with new ones.',
          inputSchema: {
            type: 'object',
            properties: {
              email_id: { type: 'string', description: 'Gmail message ID' },
              label_ids: { type: 'array', items: { type: 'string' }, description: 'Array of label IDs to set' },
            },
            required: ['email_id', 'label_ids'],
          },
        },
        {
          name: 'get_email_labels',
          description: 'Get all labels on a specific email.',
          inputSchema: {
            type: 'object',
            properties: {
              email_id: { type: 'string', description: 'Gmail message ID' },
            },
            required: ['email_id'],
          },
        },
        {
          name: 'search_by_label',
          description: 'Search emails by label(s).',
          inputSchema: {
            type: 'object',
            properties: {
              label_names: { type: 'array', items: { type: 'string' }, description: 'Array of label names to search for' },
              label_ids: { type: 'array', items: { type: 'string' }, description: 'Array of label IDs to search for' },
              additional_query: { type: 'string', description: 'Additional Gmail search query to combine with label search' },
              max: { type: 'number', description: 'Max emails to return (default 50)', default: 50 },
            },
          },
        },
        {
          name: 'bulk_label_emails',
          description: 'High-performance bulk labeling operations for large numbers of emails.',
          inputSchema: {
            type: 'object',
            properties: {
              query: { type: 'string', description: 'Gmail search query to find emails to label' },
              label_ids: { type: 'array', items: { type: 'string' }, description: 'Array of label IDs to add' },
              operation: { type: 'string', description: 'Operation type: add, remove, or replace', default: 'add' },
              batch_size: { type: 'number', description: 'Batch size for processing (default 100)', default: 100 },
              dry_run: { type: 'boolean', description: 'Show what would be labeled without actually labeling', default: false },
            },
            required: ['query', 'label_ids'],
          },
        },
        {
          name: 'auto_label_emails',
          description: 'Automatically label emails based on rules (sender, subject patterns).',
          inputSchema: {
            type: 'object',
            properties: {
              rules: { 
                type: 'array', 
                items: { 
                  type: 'object',
                  properties: {
                    label_name: { type: 'string', description: 'Label name to apply' },
                    sender_pattern: { type: 'string', description: 'Email sender pattern (e.g., "@company.com")' },
                    subject_pattern: { type: 'string', description: 'Subject pattern (e.g., "urgent")' },
                    query: { type: 'string', description: 'Custom Gmail search query' },
                  },
                  required: ['label_name']
                },
                description: 'Array of labeling rules' 
              },
              dry_run: { type: 'boolean', description: 'Show what would be labeled without actually labeling', default: false },
            },
            required: ['rules'],
          },
        },
        {
          name: 'label_emails_by_query',
          description: 'Label emails matching a search query.',
          inputSchema: {
            type: 'object',
            properties: {
              query: { type: 'string', description: 'Gmail search query to find emails to label' },
              label_name: { type: 'string', description: 'Label name to apply' },
              label_id: { type: 'string', description: 'Label ID to apply (alternative to label_name)' },
              max: { type: 'number', description: 'Max emails to process (default 100)', default: 100 },
              dry_run: { type: 'boolean', description: 'Show what would be labeled without actually labeling', default: false },
            },
            required: ['query'],
          },
        },
        {
          name: 'get_label_statistics',
          description: 'Get statistics about label usage.',
          inputSchema: {
            type: 'object',
            properties: {
              label_ids: { type: 'array', items: { type: 'string' }, description: 'Specific label IDs to get stats for (optional)' },
            },
          },
        },
        {
          name: 'cleanup_unused_labels',
          description: 'Remove labels that have no emails.',
          inputSchema: {
            type: 'object',
            properties: {
              dry_run: { type: 'boolean', description: 'Show what would be deleted without actually deleting', default: true },
            },
          },
        },
        // Auto-Labeling Rules Management
        {
          name: 'list_auto_labeling_rules',
          description: 'List all stored auto-labeling rules.',
          inputSchema: {
            type: 'object',
            properties: {},
          },
        },
        {
          name: 'add_auto_labeling_rule',
          description: 'Add a new auto-labeling rule to the configuration.',
          inputSchema: {
            type: 'object',
            properties: {
              label_name: { type: 'string', description: 'Label name to apply' },
              sender_pattern: { type: 'string', description: 'Email sender pattern (e.g., "@company.com")' },
              subject_pattern: { type: 'string', description: 'Subject pattern (e.g., "urgent")' },
              subject_contains: { type: 'array', items: { type: 'string' }, description: 'Array of strings that subject must contain (OR logic)' },
              query: { type: 'string', description: 'Custom Gmail search query' },
              enabled: { type: 'boolean', description: 'Whether the rule is enabled', default: true },
            },
            required: ['label_name'],
          },
        },
        {
          name: 'remove_auto_labeling_rule',
          description: 'Remove an auto-labeling rule by index.',
          inputSchema: {
            type: 'object',
            properties: {
              rule_index: { type: 'number', description: 'Index of the rule to remove (0-based)' },
            },
            required: ['rule_index'],
          },
        },
        {
          name: 'update_auto_labeling_rule',
          description: 'Update an existing auto-labeling rule.',
          inputSchema: {
            type: 'object',
            properties: {
              rule_index: { type: 'number', description: 'Index of the rule to update (0-based)' },
              label_name: { type: 'string', description: 'Label name to apply' },
              sender_pattern: { type: 'string', description: 'Email sender pattern (e.g., "@company.com")' },
              subject_pattern: { type: 'string', description: 'Subject pattern (e.g., "urgent")' },
              subject_contains: { type: 'array', items: { type: 'string' }, description: 'Array of strings that subject must contain (OR logic)' },
              query: { type: 'string', description: 'Custom Gmail search query' },
              enabled: { type: 'boolean', description: 'Whether the rule is enabled' },
            },
            required: ['rule_index'],
          },
        },
        {
          name: 'run_auto_labeling_rules',
          description: 'Run all enabled auto-labeling rules on emails with pagination support.',
          inputSchema: {
            type: 'object',
            properties: {
              dry_run: { type: 'boolean', description: 'Show what would be labeled without actually labeling', default: false },
              max_per_rule: { type: 'number', description: 'Maximum emails to process per rule (default 100)', default: 100 },
              batch_size: { type: 'number', description: 'Number of emails per batch (default 100)', default: 100 },
              max_batches: { type: 'number', description: 'Maximum number of batches to process (default 10)', default: 10 },
            },
          },
        },
        {
          name: 'export_auto_labeling_rules',
          description: 'Export auto-labeling rules to a JSON file.',
          inputSchema: {
            type: 'object',
            properties: {
              file_path: { type: 'string', description: 'Path to save the exported rules (optional)' },
            },
          },
        },
        {
          name: 'import_auto_labeling_rules',
          description: 'Import auto-labeling rules from a JSON file.',
          inputSchema: {
            type: 'object',
            properties: {
              file_path: { type: 'string', description: 'Path to the JSON file to import' },
              merge: { type: 'boolean', description: 'Merge with existing rules instead of replacing', default: false },
            },
            required: ['file_path'],
          },
        },
      ],
    }));

    this.server.setRequestHandler(CallToolRequestSchema, async (request) => {
      const { name, arguments: args } = request.params;
      try {
        switch (name) {
          case 'start_oauth':
            return await this.toolStartOAuth(args);
          case 'auth_status':
            return await this.toolAuthStatus();
          case 'list_unread':
            return await this.toolListUnread(args);
          case 'list_recent_unread':
            return await this.toolListRecentUnread(args);
          case 'get_message':
            return await this.toolGetMessage(args);
          case 'reply_to_message':
            return await this.toolReplyToMessage(args);
          case 'mark_as_read':
            return await this.toolMarkAsRead(args);
          case 'batch_archive':
            return await this.toolBatchArchive(args);
          case 'batch_delete':
            return await this.toolBatchDelete(args);
          case 'start_ui':
            return await this.toolStartUi(args);
          case 'stop_ui':
            return await this.toolStopUi(args);
          case 'search_emails':
            return await this.toolSearchEmails(args);
          case 'delete_email':
            return await this.toolDeleteEmail(args);
          case 'delete_emails_by_query':
            return await this.toolDeleteEmailsByQuery(args);
          case 'bulk_delete_emails':
            return await this.toolBulkDeleteEmails(args);
          case 'create_draft':
            return await this.toolCreateDraft(args);
          case 'create_reply_draft':
            return await this.toolCreateReplyDraft(args);
          case 'get_draft':
            return await this.toolGetDraft(args);
          case 'list_drafts':
            return await this.toolListDrafts(args);
          case 'update_draft':
            return await this.toolUpdateDraft(args);
          case 'send_draft':
            return await this.toolSendDraft(args);
          case 'delete_draft':
            return await this.toolDeleteDraft(args);
          case 'snooze_email':
            return await this.toolSnoozeEmail(args);
          case 'unsnooze_email':
            return await this.toolUnsnoozeEmail(args);
          case 'list_snoozed_emails':
            return await this.toolListSnoozedEmails(args);
          // Label Management Tools
          case 'create_label':
            return await this.toolCreateLabel(args);
          case 'create_labels':
            return await this.toolCreateLabels(args);
          case 'list_labels':
            return await this.toolListLabels(args);
          case 'get_label':
            return await this.toolGetLabel(args);
          case 'update_label':
            return await this.toolUpdateLabel(args);
          case 'delete_label':
            return await this.toolDeleteLabel(args);
          case 'label_email':
            return await this.toolLabelEmail(args);
          case 'label_emails':
            return await this.toolLabelEmails(args);
          case 'unlabel_email':
            return await this.toolUnlabelEmail(args);
          case 'unlabel_emails':
            return await this.toolUnlabelEmails(args);
          case 'set_email_labels':
            return await this.toolSetEmailLabels(args);
          case 'get_email_labels':
            return await this.toolGetEmailLabels(args);
          case 'search_by_label':
            return await this.toolSearchByLabel(args);
          case 'bulk_label_emails':
            return await this.toolBulkLabelEmails(args);
          case 'auto_label_emails':
            return await this.toolAutoLabelEmails(args);
          case 'label_emails_by_query':
            return await this.toolLabelEmailsByQuery(args);
          case 'get_label_statistics':
            return await this.toolGetLabelStatistics(args);
          case 'cleanup_unused_labels':
            return await this.toolCleanupUnusedLabels(args);
          // Auto-Labeling Rules Management
          case 'list_auto_labeling_rules':
            return await this.toolListAutoLabelingRules(args);
          case 'add_auto_labeling_rule':
            return await this.toolAddAutoLabelingRule(args);
          case 'remove_auto_labeling_rule':
            return await this.toolRemoveAutoLabelingRule(args);
          case 'update_auto_labeling_rule':
            return await this.toolUpdateAutoLabelingRule(args);
          case 'run_auto_labeling_rules':
            return await this.toolRunAutoLabelingRules(args);
          case 'export_auto_labeling_rules':
            return await this.toolExportAutoLabelingRules(args);
          case 'import_auto_labeling_rules':
            return await this.toolImportAutoLabelingRules(args);
          default:
            throw new McpError(ErrorCode.MethodNotFound, `Tool ${name} not found`);
        }
      } catch (error) {
        if (error instanceof McpError) throw error;
        throw new McpError(ErrorCode.InternalError, error.message || 'Unknown error');
      }
    });
  }

  async toolStartOAuth(args) {
    const port = args?.port || this.oauthState.redirectPort;
    this.oauthState.redirectPort = port;
    const oAuth2Client = await this.getOAuth2Client();
    await this.startLocalOAuthServer(oAuth2Client, port);
    const authUrl = oAuth2Client.generateAuthUrl({
      access_type: 'offline',
      scope: SCOPES,
      prompt: 'consent',
    });
    return {
      content: this.formatContent([
        'Open this URL to authorize Gmail access:',
        authUrl,
        `Listening on http://127.0.0.1:${port}/oauth2callback for the redirect...`,
      ]),
      isText: true,
    };
  }

  async toolAuthStatus() {
    const token = await this.loadToken();
    if (!token) {
      return { content: this.formatContent('Not authorized'), isText: true };
    }
    try {
      const gmail = await this.getGmail();
      const profile = await gmail.users.getProfile({ userId: 'me' });
      return {
        content: this.formatContent(`Authorized as ${profile.data.emailAddress}`),
        isText: true,
      };
    } catch {
      return { content: this.formatContent('Authorized, but failed to fetch profile'), isText: true };
    }
  }

  headerValue(headers, name) {
    const h = headers?.find(h => h.name?.toLowerCase() === name.toLowerCase());
    return h?.value || '';
  }

  async toolListUnread(args) {
    const max = Math.max(1, Math.min(50, Number(args?.max || 10)));
    const gmail = await this.getGmail();
    const list = await gmail.users.messages.list({ userId: 'me', q: 'is:unread', maxResults: max });
    const messages = list.data.messages || [];
    if (messages.length === 0) {
      return { content: this.formatContent('No unread emails.'), isText: true };
    }

    const results = [];
    for (const m of messages) {
      const msg = await gmail.users.messages.get({ userId: 'me', id: m.id, format: 'metadata', metadataHeaders: ['From','Subject','Date','Message-ID'] });
      const headers = msg.data.payload?.headers || [];
      const from = this.headerValue(headers, 'From');
      const subject = this.headerValue(headers, 'Subject');
      const date = this.headerValue(headers, 'Date');
      const snippet = msg.data.snippet || '';
      results.push({ id: m.id, threadId: msg.data.threadId, from, subject, date, snippet });
    }

    return { content: this.formatContent(results), isText: true };
  }

  parseDaysParameter(daysStr) {
    if (!daysStr) return 7;
    
    // Handle various formats: "2", "3d", "2 days", "--2", "2d"
    const str = String(daysStr).toLowerCase().trim();
    
    // Remove common prefixes
    const cleaned = str.replace(/^--/, '').replace(/^days?\s*/, '');
    
    // Extract number
    const match = cleaned.match(/^(\d+)/);
    if (match) {
      return Math.max(1, Math.min(30, parseInt(match[1], 10))); // Limit to 1-30 days
    }
    
    return 7; // Default fallback
  }

  formatEmailsAsTable(emails) {
    if (!emails || emails.length === 0) {
      return "No unread emails found.";
    }

    // Create table header
    let table = "| ID | Date | Sender | Subject |\n";
    table += "|---|---|---|---|\n";

    // Add each email as a table row
    emails.forEach((email, index) => {
      const id = index + 1;
      
      // Format date as DD/MM HH:MM
      const date = new Date(email.date);
      const day = String(date.getDate()).padStart(2, '0');
      const month = String(date.getMonth() + 1).padStart(2, '0');
      const hours = String(date.getHours()).padStart(2, '0');
      const minutes = String(date.getMinutes()).padStart(2, '0');
      const formattedDate = `${day}/${month} ${hours}:${minutes}`;
      
      // Extract sender name (before <email>)
      const sender = email.from.split('<')[0].trim();
      
      // Truncate long subjects
      const subject = email.subject.length > 50 
        ? email.subject.substring(0, 47) + '...' 
        : email.subject;
      
      table += `| ${id} | ${formattedDate} | ${sender} | ${subject} |\n`;
    });

    // Add summary
    table += `\n**Total: ${emails.length} unread emails**\n`;
    
    return table;
  }

  async toolListRecentUnread(args) {
    const days = this.parseDaysParameter(args?.days);
    const max = Math.max(1, Math.min(100, Number(args?.max || 50)));
    const gmail = await this.getGmail();
    
    // Load ignored labels configuration
    const ignoredLabels = await this.loadIgnoredLabels();
    
    // Build Gmail search query for recent unread emails from Inbox only
    let query = `in:inbox is:unread newer_than:${days}d`;
    
    // Add exclusions for ignored labels
    if (ignoredLabels.length > 0) {
      const exclusionQueries = ignoredLabels.map(label => `-label:"${label}"`);
      query += ' ' + exclusionQueries.join(' ');
    }
    
    const list = await gmail.users.messages.list({ userId: 'me', q: query, maxResults: max });
    const messages = list.data.messages || [];
    
    if (messages.length === 0) {
      const ignoredInfo = ignoredLabels.length > 0 ? ` (excluding ${ignoredLabels.length} ignored labels)` : '';
      return { content: this.formatContent(`No unread emails from the past ${days} days in Inbox${ignoredInfo}.`), isText: true };
    }

    const results = [];
    for (const m of messages) {
      const msg = await gmail.users.messages.get({ userId: 'me', id: m.id, format: 'metadata', metadataHeaders: ['From','Subject','Date','Message-ID'] });
      const headers = msg.data.payload?.headers || [];
      const from = this.headerValue(headers, 'From');
      const subject = this.headerValue(headers, 'Subject');
      const date = this.headerValue(headers, 'Date');
      const snippet = msg.data.snippet || '';
      results.push({ id: m.id, threadId: msg.data.threadId, from, subject, date, snippet });
    }

    // Format as table
    const tableOutput = this.formatEmailsAsTable(results);
    return { content: this.formatContent(tableOutput), isText: true };
  }

  extractPlainText(payload) {
    if (!payload) return '';
    if (payload.mimeType === 'text/plain' && payload.body?.data) {
      return decodeBase64Url(payload.body.data);
    }
    if (payload.parts && Array.isArray(payload.parts)) {
      for (const part of payload.parts) {
        const text = this.extractPlainText(part);
        if (text) return text;
      }
    }
    return '';
  }

  extractHtmlContent(payload) {
    if (!payload) return '';
    if (payload.mimeType === 'text/html' && payload.body?.data) {
      return decodeBase64Url(payload.body.data);
    }
    if (payload.parts && Array.isArray(payload.parts)) {
      for (const part of payload.parts) {
        const html = this.extractHtmlContent(part);
        if (html) return html;
      }
    }
    return '';
  }

  extractEmailContent(payload) {
    const plainText = this.extractPlainText(payload);
    const htmlContent = this.extractHtmlContent(payload);
    
    return {
      plainText,
      htmlContent,
      hasHtml: !!htmlContent
    };
  }

  async toolGetMessage(args) {
    if (!args?.id) throw new McpError(ErrorCode.InvalidParams, 'id is required');
    const gmail = await this.getGmail();
    const msg = await gmail.users.messages.get({ userId: 'me', id: args.id, format: 'full' });
    const headers = msg.data.payload?.headers || [];
    const from = this.headerValue(headers, 'From');
    const to = this.headerValue(headers, 'To');
    const subject = this.headerValue(headers, 'Subject');
    const date = this.headerValue(headers, 'Date');
    const messageIdHeader = this.headerValue(headers, 'Message-ID');
    const body = this.extractPlainText(msg.data.payload) || msg.data.snippet || '';

    const result = { id: msg.data.id, threadId: msg.data.threadId, from, to, subject, date, messageIdHeader, snippet: msg.data.snippet, body };
    return { content: this.formatContent(result), isText: true };
  }

  async toolReplyToMessage(args) {
    const { message_id, body } = args || {};
    if (!message_id || !body) throw new McpError(ErrorCode.InvalidParams, 'message_id and body are required');

    const gmail = await this.getGmail();
    const original = await gmail.users.messages.get({ userId: 'me', id: message_id, format: 'full' });
    const headers = original.data.payload?.headers || [];
    const from = this.headerValue(headers, 'From');
    const subject = this.headerValue(headers, 'Subject');
    const messageIdHeader = this.headerValue(headers, 'Message-ID');
    const threadId = original.data.threadId;

    const replySubject = subject?.toLowerCase().startsWith('re:') ? subject : `Re: ${subject}`;

    const raw = [
      'Content-Type: text/plain; charset="UTF-8"',
      'MIME-Version: 1.0',
      `To: ${from}`,
      `Subject: ${replySubject}`,
      messageIdHeader ? `In-Reply-To: ${messageIdHeader}` : null,
      messageIdHeader ? `References: ${messageIdHeader}` : null,
      '',
      body,
    ].filter(Boolean).join('\r\n');

    const encodedMessage = base64UrlEncode(raw);

    const sendRes = await gmail.users.messages.send({
      userId: 'me',
      requestBody: { raw: encodedMessage, threadId },
    });

    // Automatically show recent unread emails table after replying
    const recentUnreadResult = await this.toolListRecentUnread({ days: '7', max: 50 });
    return { 
      content: this.formatContent([
        `Reply sent. New message ID: ${sendRes.data.id}`,
        '',
        'Recent unread emails:',
        recentUnreadResult.content[0].text
      ]), 
      isText: true 
    };
  }

  async toolMarkAsRead(args) {
    const { id } = args || {};
    if (!id) throw new McpError(ErrorCode.InvalidParams, 'id is required');
    const gmail = await this.getGmail();
    await gmail.users.messages.modify({ userId: 'me', id, requestBody: { removeLabelIds: ['UNREAD'] } });
    
    // Automatically show recent unread emails table after marking as read
    const recentUnreadResult = await this.toolListRecentUnread({ days: '7', max: 50 });
    return { 
      content: this.formatContent([
        'Message marked as read.',
        '',
        'Recent unread emails:',
        recentUnreadResult.content[0].text
      ]), 
      isText: true 
    };
  }

  async toolBatchArchive(args) {
    const ids = Array.isArray(args?.ids) ? args.ids : [];
    if (ids.length === 0) throw new McpError(ErrorCode.InvalidParams, 'ids is required (non-empty array)');
    const gmail = await this.getGmail();
    await gmail.users.messages.batchModify({ userId: 'me', requestBody: { ids, removeLabelIds: ['INBOX', 'UNREAD'] } });
    
    // Automatically show recent unread emails table after archiving
    const recentUnreadResult = await this.toolListRecentUnread({ days: '7', max: 50 });
    return { 
      content: this.formatContent([
        `Archived ${ids.length} messages`,
        '',
        'Recent unread emails:',
        recentUnreadResult.content[0].text
      ]), 
      isText: true 
    };
  }

  async toolBatchDelete(args) {
    const ids = Array.isArray(args?.ids) ? args.ids : [];
    const permanent = !!args?.permanent;
    if (ids.length === 0) throw new McpError(ErrorCode.InvalidParams, 'ids is required (non-empty array)');
    const gmail = await this.getGmail();
    
    let actionMessage;
    if (permanent) {
      // Permanent delete (be careful!)
      await gmail.users.messages.batchDelete({ userId: 'me', requestBody: { ids } });
      actionMessage = `Permanently deleted ${ids.length} messages`;
    } else {
      // Safer: move to Trash
      await Promise.all(ids.map(id => gmail.users.messages.trash({ userId: 'me', id })));
      actionMessage = `Moved ${ids.length} messages to Trash`;
    }
    
    // Automatically show recent unread emails table after deleting
    const recentUnreadResult = await this.toolListRecentUnread({ days: '7', max: 50 });
    return { 
      content: this.formatContent([
        actionMessage,
        '',
        'Recent unread emails:',
        recentUnreadResult.content[0].text
      ]), 
      isText: true 
    };
  }

  async toolStartUi(args) {
    const port = Number(args?.port || this.uiState.port);
    const query = String(args?.query || 'is:unread');
    const max = Math.max(1, Math.min(100, Number(args?.max || 25)));
    const gmail = await this.getGmail();

    // Clean up previous server instance
    if (this.uiState.server) {
      try { 
        this.uiState.server.close(); 
        // Wait a bit for the server to actually close
        await new Promise(resolve => setTimeout(resolve, 100));
      } catch {}
      this.uiState.server = null;
    }

    const sendJson = (res, obj) => {
      const body = JSON.stringify(obj);
      res.writeHead(200, { 'Content-Type': 'application/json' });
      res.end(body);
    };

    const parseBody = async (req) => new Promise((resolve, reject) => {
      let data = '';
      req.on('data', chunk => {
        data += chunk;
        if (data.length > 5 * 1024 * 1024) { // 5MB guard
          reject(new Error('payload too large'));
          req.destroy();
        }
      });
      req.on('end', () => {
        try { resolve(data ? JSON.parse(data) : {}); } catch (e) { reject(e); }
      });
      req.on('error', reject);
    });

    const server = http.createServer(async (req, res) => {
      try {
        const url = new URL(req.url, `http://127.0.0.1:${port}`);
        
        if (req.method === 'GET' && url.pathname === '/') {
          const html = `<!DOCTYPE html>
<html lang="en">
  <head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Gmail MCP UI</title>
    <style>
      * { margin: 0; padding: 0; box-sizing: border-box; }
      body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; background: #f5f5f5; }
      .container { max-width: 1200px; margin: 0 auto; padding: 20px; }
      .header { background: white; padding: 20px; border-radius: 8px; margin-bottom: 20px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
      .search-bar { width: 100%; padding: 12px; border: 1px solid #ddd; border-radius: 6px; font-size: 16px; }
      .email-list { background: white; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
      .email-item { padding: 16px; border-bottom: 1px solid #eee; cursor: pointer; display: flex; align-items: center; gap: 12px; }
      .email-item:hover { background: #f9f9f9; }
      .email-item.selected { background: #e3f2fd; }
      .email-checkbox { margin-right: 8px; }
      .email-content { flex: 1; }
      .email-header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 4px; }
      .email-from { font-weight: 600; color: #333; }
      .email-date { color: #666; font-size: 14px; }
      .email-subject { font-weight: 500; margin-bottom: 4px; }
      .email-snippet { color: #666; font-size: 14px; }
      .actions { background: white; padding: 16px; border-radius: 8px; margin-bottom: 20px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
      .btn { padding: 8px 16px; border: none; border-radius: 4px; cursor: pointer; margin-right: 8px; }
      .btn-primary { background: #1976d2; color: white; }
      .btn-danger { background: #d32f2f; color: white; }
      .btn-secondary { background: #757575; color: white; }
      .notification { position: fixed; top: 20px; right: 20px; padding: 12px 20px; border-radius: 4px; color: white; z-index: 1000; }
      .notification.success { background: #4caf50; }
      .notification.error { background: #f44336; }
      .loading { text-align: center; padding: 40px; color: #666; }
      .empty-state { text-align: center; padding: 40px; color: #666; }
      .empty-state-icon { font-size: 48px; margin-bottom: 16px; }
      .modal { display: none; position: fixed; z-index: 1000; left: 0; top: 0; width: 100%; height: 100%; background-color: rgba(0,0,0,0.5); }
      .modal-content { background-color: white; margin: 5% auto; padding: 20px; border-radius: 8px; width: 80%; max-width: 600px; }
      .modal-header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 20px; }
      .modal-body textarea { width: 100%; height: 200px; padding: 12px; border: 1px solid #ddd; border-radius: 4px; font-family: inherit; }
      .modal-footer { margin-top: 20px; text-align: right; }
      .close { font-size: 28px; font-weight: bold; cursor: pointer; }
    </style>
  </head>
  <body>
    <div class="container">
      <div class="header">
        <h1>Gmail MCP UI</h1>
        <input type="text" id="searchInput" class="search-bar" placeholder="Search emails..." value="is:unread">
      </div>
      
      <div class="actions">
        <button id="checkAll" type="checkbox">Select All</button>
        <button id="archive" class="btn btn-secondary">Archive</button>
        <button id="trash" class="btn btn-danger">Trash</button>
        <button id="markRead" class="btn btn-primary">Mark as Read</button>
        <button id="reply" class="btn btn-primary">Reply</button>
      </div>
      
      <div class="email-list" id="emailList">
        <div class="loading">Loading emails...</div>
      </div>
    </div>

    <!-- Reply Modal -->
    <div id="replyModal" class="modal">
      <div class="modal-content">
        <div class="modal-header">
          <h3>Reply to Email</h3>
          <span class="close">&times;</span>
        </div>
        <div class="modal-body">
          <div id="replyEmailInfo" style="margin-bottom: 15px; padding: 10px; background: #f5f5f5; border-radius: 4px;"></div>
          <textarea id="replyBody" placeholder="Type your reply here..."></textarea>
        </div>
        <div class="modal-footer">
          <button id="sendReply" class="btn btn-primary">Send Reply</button>
          <button id="cancelReply" class="btn btn-secondary">Cancel</button>
        </div>
      </div>
    </div>

    <script>
      const state = { items: [], searchQuery: 'is:unread' };

      async function load() {
        try {
          const query = document.getElementById('searchInput').value.trim() || state.searchQuery;
          const res = await fetch(\`/api/list?q=\${encodeURIComponent(query)}\`);
        const data = await res.json();
        state.items = data.items || [];
          
          const emailList = document.getElementById('emailList');
          emailList.innerHTML = '';
          
          if (state.items.length === 0) {
            emailList.innerHTML = '<div class="empty-state"><div class="empty-state-icon">📭</div><h3>No emails found</h3></div>';
          } else {
            for (const email of state.items) {
              const emailItem = document.createElement('div');
              emailItem.className = 'email-item';
              emailItem.innerHTML = \`
                <input type="checkbox" class="email-checkbox" data-id="\${email.id || ''}" />
                <div class="email-content">
                  <div class="email-header">
                    <span class="email-from">\${email.from || 'Unknown'}</span>
                    <span class="email-date">\${new Date(email.date || '').toLocaleDateString()}</span>
                  </div>
                  <div class="email-subject">\${email.subject || 'No Subject'}</div>
                  <div class="email-snippet">\${email.snippet || ''}</div>
                </div>
              \`;
              emailList.appendChild(emailItem);
            }
          }
        } catch (error) {
          console.error('Error loading messages:', error);
          document.getElementById('emailList').innerHTML = '<div class="empty-state"><div class="empty-state-icon">❌</div><h3>Error loading emails</h3></div>';
        }
      }

      document.getElementById('searchInput').addEventListener('keypress', (e) => {
        if (e.key === 'Enter') load();
      });

      document.getElementById('checkAll').addEventListener('change', (e) => {
        document.querySelectorAll('.email-checkbox').forEach(cb => cb.checked = e.target.checked);
      });

      document.getElementById('archive').addEventListener('click', async () => {
        const ids = Array.from(document.querySelectorAll('.email-checkbox:checked')).map(cb => cb.getAttribute('data-id'));
        if (!ids.length) return;
        try {
          await fetch('/api/archive', { method: 'POST', headers: {'Content-Type': 'application/json'}, body: JSON.stringify({ids}) });
          load();
        } catch (error) {
          console.error('Archive error:', error);
        }
      });

      document.getElementById('trash').addEventListener('click', async () => {
        const ids = Array.from(document.querySelectorAll('.email-checkbox:checked')).map(cb => cb.getAttribute('data-id'));
        if (!ids.length) return;
        try {
          await fetch('/api/delete', { method: 'POST', headers: {'Content-Type': 'application/json'}, body: JSON.stringify({ids}) });
          load();
        } catch (error) {
          console.error('Trash error:', error);
        }
      });

      document.getElementById('markRead').addEventListener('click', async () => {
        const ids = Array.from(document.querySelectorAll('.email-checkbox:checked')).map(cb => cb.getAttribute('data-id'));
        if (!ids.length) return;
        try {
          await Promise.all(ids.map(id => fetch('/api/markRead', { method: 'POST', headers: {'Content-Type': 'application/json'}, body: JSON.stringify({id}) })));
          load();
        } catch (error) {
          console.error('Mark read error:', error);
        }
      });

      // Reply functionality
      let selectedEmailForReply = null;

      document.getElementById('reply').addEventListener('click', async () => {
        const checkedBoxes = document.querySelectorAll('.email-checkbox:checked');
        if (checkedBoxes.length === 0) {
          alert('Please select an email to reply to');
          return;
        }
        if (checkedBoxes.length > 1) {
          alert('Please select only one email to reply to');
          return;
        }

        const emailId = checkedBoxes[0].getAttribute('data-id');
        const emailItem = checkedBoxes[0].closest('.email-item');
        const from = emailItem.querySelector('.email-from').textContent;
        const subject = emailItem.querySelector('.email-subject').textContent;

        selectedEmailForReply = emailId;
        document.getElementById('replyEmailInfo').innerHTML = \`
          <strong>To:</strong> \${from}<br>
          <strong>Subject:</strong> \${subject}
        \`;
        document.getElementById('replyBody').value = '';
        document.getElementById('replyModal').style.display = 'block';
      });

      document.getElementById('sendReply').addEventListener('click', async () => {
        const body = document.getElementById('replyBody').value.trim();
        if (!body) {
          alert('Please enter a reply message');
          return;
        }
        if (!selectedEmailForReply) {
          alert('No email selected for reply');
          return;
        }

        try {
          await fetch('/api/reply', { 
            method: 'POST', 
            headers: {'Content-Type': 'application/json'}, 
            body: JSON.stringify({
              message_id: selectedEmailForReply,
              body: body
            }) 
          });
          document.getElementById('replyModal').style.display = 'none';
          load(); // Refresh the email list
        } catch (error) {
          console.error('Reply error:', error);
          alert('Error sending reply');
        }
      });

      document.getElementById('cancelReply').addEventListener('click', () => {
        document.getElementById('replyModal').style.display = 'none';
        selectedEmailForReply = null;
      });

      document.querySelector('.close').addEventListener('click', () => {
        document.getElementById('replyModal').style.display = 'none';
        selectedEmailForReply = null;
      });

      // Close modal when clicking outside
      window.addEventListener('click', (event) => {
        const modal = document.getElementById('replyModal');
        if (event.target === modal) {
          modal.style.display = 'none';
          selectedEmailForReply = null;
        }
      });

      load();
    </script>
  </body>
</html>`;
          res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' });
          res.end(html);
          return;
        }

        if (req.method === 'GET' && url.pathname === '/api/list') {
          const maxParam = Number(url.searchParams.get('max') || max);
          const q = url.searchParams.get('q') || query;
          const list = await gmail.users.messages.list({ userId: 'me', q, maxResults: Math.max(1, Math.min(100, maxParam)) });
          const messages = list.data.messages || [];
          const items = [];
          for (const m of messages) {
            const msg = await gmail.users.messages.get({ userId: 'me', id: m.id, format: 'metadata', metadataHeaders: ['From','Subject','Date'] });
            const headers = msg.data.payload?.headers || [];
            const hv = (n) => (headers.find(h => h.name?.toLowerCase() === n.toLowerCase())?.value || '');
            const labels = msg.data.labelIds || [];
            const unread = labels.includes('UNREAD');
            items.push({ 
              id: m.id, 
              threadId: msg.data.threadId, 
              from: hv('From'), 
              subject: hv('Subject'), 
              date: hv('Date'),
              snippet: msg.data.snippet || '',
              unread: unread
            });
          }
          return sendJson(res, { items });
        }

        if (req.method === 'GET' && url.pathname.startsWith('/api/message/')) {
          const messageId = url.pathname.split('/')[3];
          if (!messageId) { res.writeHead(400); res.end('Bad request'); return; }
          const msg = await gmail.users.messages.get({ userId: 'me', id: messageId, format: 'full' });
          const headers = msg.data.payload?.headers || [];
          const hv = (n) => (headers.find(h => h.name?.toLowerCase() === n.toLowerCase())?.value || '');
          const emailContent = this.extractEmailContent(msg.data.payload);
          const result = { 
            id: msg.data.id, 
            threadId: msg.data.threadId, 
            from: hv('From'), 
            to: hv('To'),
            subject: hv('Subject'), 
            date: hv('Date'),
            snippet: msg.data.snippet,
            body: emailContent.plainText,
            htmlBody: emailContent.htmlContent,
            hasHtml: emailContent.hasHtml
          };
          return sendJson(res, result);
        }

        if (req.method === 'POST' && url.pathname === '/api/archive') {
          const body = await parseBody(req);
          const ids = Array.isArray(body?.ids) ? body.ids : [];
          await gmail.users.messages.batchModify({ userId: 'me', requestBody: { ids, removeLabelIds: ['INBOX', 'UNREAD'] } });
          return sendJson(res, { ok: true });
        }

        if (req.method === 'POST' && url.pathname === '/api/delete') {
          const body = await parseBody(req);
          const ids = Array.isArray(body?.ids) ? body.ids : [];
          await Promise.all(ids.map(id => gmail.users.messages.trash({ userId: 'me', id })));
          return sendJson(res, { ok: true });
        }

        if (req.method === 'POST' && url.pathname === '/api/reply') {
          const body = await parseBody(req);
          const message_id = body?.message_id;
          const text = body?.body || '';
          const html = body?.html || '';
          const useHtml = body?.useHtml || false;
          if (!message_id || (!text && !html)) { res.writeHead(400); res.end('Bad request'); return; }
          
          const original = await gmail.users.messages.get({ userId: 'me', id: message_id, format: 'full' });
          const headers = original.data.payload?.headers || [];
          const hv = (n) => (headers.find(h => h.name?.toLowerCase() === n.toLowerCase())?.value || '');
          const from = hv('From');
          const subject = hv('Subject');
          const messageIdHeader = hv('Message-ID');
          const replySubject = subject?.toLowerCase().startsWith('re:') ? subject : `Re: ${subject}`;
          
          let raw;
          if (useHtml && html) {
            // HTML email with plain text fallback
            const boundary = '----=_Part_' + Math.random().toString(36).substr(2, 9);
            raw = [
              `To: ${from}`,
              `Subject: ${replySubject}`,
              messageIdHeader ? `In-Reply-To: ${messageIdHeader}` : null,
              messageIdHeader ? `References: ${messageIdHeader}` : null,
              'MIME-Version: 1.0',
              `Content-Type: multipart/alternative; boundary="${boundary}"`,
              '',
              `--${boundary}`,
              'Content-Type: text/plain; charset="UTF-8"',
              'Content-Transfer-Encoding: 7bit',
              '',
              text || html.replace(/<[^>]*>/g, ''), // Fallback to plain text
              '',
              `--${boundary}`,
              'Content-Type: text/html; charset="UTF-8"',
              'Content-Transfer-Encoding: 7bit',
              '',
              html,
              '',
              `--${boundary}--`
            ].filter(Boolean).join('\r\n');
          } else {
            // Plain text email
            raw = [
            'Content-Type: text/plain; charset="UTF-8"',
            'MIME-Version: 1.0',
            `To: ${from}`,
            `Subject: ${replySubject}`,
            messageIdHeader ? `In-Reply-To: ${messageIdHeader}` : null,
            messageIdHeader ? `References: ${messageIdHeader}` : null,
            '',
            text,
          ].filter(Boolean).join('\r\n');
          }
          
          const encoded = Buffer.from(raw).toString('base64').replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
          await gmail.users.messages.send({ userId: 'me', requestBody: { raw: encoded, threadId: original.data.threadId } });
          return sendJson(res, { ok: true });
        }

        if (req.method === 'POST' && url.pathname === '/api/markRead') {
          const body = await parseBody(req);
          const id = body?.id;
          if (!id) { res.writeHead(400); res.end('Bad request'); return; }
          await gmail.users.messages.modify({ userId: 'me', id, requestBody: { removeLabelIds: ['UNREAD'] } });
          return sendJson(res, { ok: true });
        }

        res.writeHead(404);
        res.end('Not found');
      } catch (e) {
        res.writeHead(500);
        res.end('Server error');
      }
    });

    // Try to start server with better error handling
    try {
    await new Promise((resolve, reject) => {
        server.once('error', (err) => {
          if (err.code === 'EADDRINUSE') {
            // Try next available port
            const nextPort = port + 1;
            reject(new Error(`Port ${port} in use, trying ${nextPort}`));
          } else {
            reject(err);
          }
        });
      server.listen(port, '127.0.0.1', () => resolve());
    });
      
    this.uiState.server = server;
    this.uiState.port = port;

    return { content: this.formatContent(`UI running at http://127.0.0.1:${port}/`), isText: true };
    } catch (error) {
      if (error.message.includes('Port') && error.message.includes('in use')) {
        // Recursively try the next port
        return this.toolStartUi({ ...args, port: port + 1 });
      }
      throw error;
    }
  }

  async toolStopUi(args) {
    if (this.uiState.server) {
      try {
        this.uiState.server.close();
        this.uiState.server = null;
        return { content: this.formatContent('UI server stopped'), isText: true };
      } catch (error) {
        return { content: this.formatContent(`Error stopping server: ${error.message}`), isText: true };
      }
    }
    return { content: this.formatContent('No UI server running'), isText: true };
  }

  async toolSearchEmails(args) {
    const query = args?.query;
    const max = Math.max(1, Math.min(100, Number(args?.max || 10)));
    
    if (!query) {
      throw new McpError(ErrorCode.InvalidParams, 'query is required');
    }

    const gmail = await this.getGmail();
    const list = await gmail.users.messages.list({ userId: 'me', q: query, maxResults: max });
    const messages = list.data.messages || [];
    
    if (messages.length === 0) {
      return { content: this.formatContent(`No emails found for query: "${query}"`), isText: true };
    }

    const results = [];
    for (const m of messages) {
      const msg = await gmail.users.messages.get({ userId: 'me', id: m.id, format: 'metadata', metadataHeaders: ['From','Subject','Date','Message-ID'] });
      const headers = msg.data.payload?.headers || [];
      const from = this.headerValue(headers, 'From');
      const subject = this.headerValue(headers, 'Subject');
      const date = this.headerValue(headers, 'Date');
      const snippet = msg.data.snippet || '';
      const labels = msg.data.labelIds || [];
      const unread = labels.includes('UNREAD');
      
      results.push({ 
        id: m.id, 
        threadId: msg.data.threadId, 
        from, 
        subject, 
        date, 
        snippet,
        unread
      });
    }

    return { 
      content: this.formatContent({
        query: query,
        count: results.length,
        emails: results
      }), 
      isText: true 
    };
  }

  async toolDeleteEmail(args) {
    const { id, permanent } = args || {};
    if (!id) throw new McpError(ErrorCode.InvalidParams, 'id is required');
    
    const gmail = await this.getGmail();
    
    // Get email details before deleting
    const msg = await gmail.users.messages.get({ userId: 'me', id, format: 'metadata', metadataHeaders: ['From','Subject','Date'] });
    const headers = msg.data.payload?.headers || [];
    const from = this.headerValue(headers, 'From');
    const subject = this.headerValue(headers, 'Subject');
    const date = this.headerValue(headers, 'Date');
    
    let actionMessage;
    if (permanent) {
      // Permanent delete (be careful!)
      await gmail.users.messages.delete({ userId: 'me', id });
      actionMessage = `Permanently deleted email: "${subject}" from ${from} (${date})`;
    } else {
      // Safer: move to Trash
      await gmail.users.messages.trash({ userId: 'me', id });
      actionMessage = `Moved to trash: "${subject}" from ${from} (${date})`;
    }
    
    // Automatically show recent unread emails table after deleting
    const recentUnreadResult = await this.toolListRecentUnread({ days: '7', max: 50 });
    return { 
      content: this.formatContent([
        actionMessage,
        '',
        'Recent unread emails:',
        recentUnreadResult.content[0].text
      ]), 
      isText: true 
    };
  }

  async toolDeleteEmailsByQuery(args) {
    const { query, max, permanent, dry_run } = args || {};
    if (!query) throw new McpError(ErrorCode.InvalidParams, 'query is required');
    
    const maxResults = Math.max(1, Math.min(100, Number(max || 50)));
    const gmail = await this.getGmail();
    
    // Search for emails
    const list = await gmail.users.messages.list({ userId: 'me', q: query, maxResults });
    const messages = list.data.messages || [];
    
    if (messages.length === 0) {
      return { 
        content: this.formatContent(`No emails found matching query: "${query}"`), 
        isText: true 
      };
    }

    // Get email details in parallel for better performance
    const emailDetails = await Promise.all(
      messages.map(async (m) => {
        try {
          const msg = await gmail.users.messages.get({ 
            userId: 'me', 
            id: m.id, 
            format: 'metadata', 
            metadataHeaders: ['From','Subject','Date'] 
          });
          const headers = msg.data.payload?.headers || [];
          const from = this.headerValue(headers, 'From');
          const subject = this.headerValue(headers, 'Subject');
          const date = this.headerValue(headers, 'Date');
          
          return {
            id: m.id,
            from,
            subject,
            date
          };
        } catch (error) {
          return {
            id: m.id,
            from: 'Unknown',
            subject: 'Error loading',
            date: 'Unknown',
            error: error.message
          };
        }
      })
    );

    if (dry_run) {
      return { 
        content: this.formatContent({
          action: 'DRY RUN - No emails were deleted',
          query: query,
          count: emailDetails.length,
          emails: emailDetails
        }), 
        isText: true 
      };
    }

    // Delete emails using batch operations for better performance
    const deletedIds = [];
    const errors = [];
    
    if (permanent) {
      // Use batch delete for permanent deletion (much faster)
      try {
        await gmail.users.messages.batchDelete({ 
          userId: 'me', 
          requestBody: { ids: emailDetails.map(e => e.id) } 
        });
        deletedIds.push(...emailDetails.map(e => e.id));
      } catch (error) {
        // If batch fails, fall back to individual deletes
        for (const email of emailDetails) {
          try {
            await gmail.users.messages.delete({ userId: 'me', id: email.id });
            deletedIds.push(email.id);
          } catch (err) {
            errors.push({ id: email.id, error: err.message });
          }
        }
      }
    } else {
      // Use batch modify for moving to trash (much faster)
      try {
        await gmail.users.messages.batchModify({ 
          userId: 'me', 
          requestBody: { 
            ids: emailDetails.map(e => e.id),
            addLabelIds: ['TRASH']
          } 
        });
        deletedIds.push(...emailDetails.map(e => e.id));
      } catch (error) {
        // If batch fails, fall back to individual trash operations
        for (const email of emailDetails) {
          try {
            await gmail.users.messages.trash({ userId: 'me', id: email.id });
            deletedIds.push(email.id);
          } catch (err) {
            errors.push({ id: email.id, error: err.message });
          }
        }
      }
    }

    const action = permanent ? 'permanently deleted' : 'moved to trash';
    return { 
      content: this.formatContent({
        action: `${deletedIds.length} emails ${action}`,
        query: query,
        successful: deletedIds.length,
        failed: errors.length,
        errors: errors.length > 0 ? errors : undefined,
        emails: emailDetails.filter(e => deletedIds.includes(e.id))
      }), 
      isText: true 
    };
  }

  async toolBulkDeleteEmails(args) {
    const { query, batch_size, permanent, dry_run } = args || {};
    if (!query) throw new McpError(ErrorCode.InvalidParams, 'query is required');
    
    const batchSize = Math.max(10, Math.min(500, Number(batch_size || 100)));
    const gmail = await this.getGmail();
    
    // First, get total count without fetching details
    const list = await gmail.users.messages.list({ userId: 'me', q: query, maxResults: 1 });
    const totalCount = list.data.resultSizeEstimate || 0;
    
    if (totalCount === 0) {
      return { 
        content: this.formatContent(`No emails found matching query: "${query}"`), 
        isText: true 
      };
    }

    if (dry_run) {
      return { 
        content: this.formatContent({
          action: 'DRY RUN - No emails were deleted',
          query: query,
          total_estimated: totalCount,
          batch_size: batchSize,
          estimated_batches: Math.ceil(totalCount / batchSize),
          note: 'This is an estimate. Actual count may vary.'
        }), 
        isText: true 
      };
    }

    // Process in batches for better performance
    let processedCount = 0;
    let deletedCount = 0;
    let errorCount = 0;
    const errors = [];
    let nextPageToken = null;

    while (processedCount < totalCount) {
      // Get batch of message IDs
      const listParams = { 
        userId: 'me', 
        q: query, 
        maxResults: batchSize 
      };
      if (nextPageToken) {
        listParams.pageToken = nextPageToken;
      }
      
      const batchList = await gmail.users.messages.list(listParams);
      const messages = batchList.data.messages || [];
      nextPageToken = batchList.data.nextPageToken;
      
      if (messages.length === 0) break;
      
      const messageIds = messages.map(m => m.id);
      
      try {
        if (permanent) {
          // Batch permanent delete
          await gmail.users.messages.batchDelete({ 
            userId: 'me', 
            requestBody: { ids: messageIds } 
          });
        } else {
          // Batch move to trash
          await gmail.users.messages.batchModify({ 
            userId: 'me', 
            requestBody: { 
              ids: messageIds,
              addLabelIds: ['TRASH']
            } 
          });
        }
        deletedCount += messageIds.length;
      } catch (error) {
        errorCount += messageIds.length;
        errors.push({ 
          batch: Math.floor(processedCount / batchSize) + 1, 
          error: error.message,
          count: messageIds.length
        });
      }
      
      processedCount += messages.length;
      
      // Break if no more pages
      if (!nextPageToken) break;
    }

    const action = permanent ? 'permanently deleted' : 'moved to trash';
    return { 
      content: this.formatContent({
        action: `Bulk operation completed`,
        query: query,
        total_processed: processedCount,
        successful: deletedCount,
        failed: errorCount,
        batch_size: batchSize,
        batches_processed: Math.ceil(processedCount / batchSize),
        errors: errors.length > 0 ? errors : undefined,
        performance_note: 'Used batch operations for optimal speed'
      }), 
      isText: true 
    };
  }

  async toolCreateDraft(args) {
    const { to, subject, body, html, reply_to_message_id } = args || {};
    if (!to || !subject || !body) {
      throw new McpError(ErrorCode.InvalidParams, 'to, subject, and body are required');
    }

    const gmail = await this.getGmail();
    
    // Build the email content
    let raw;
    if (html) {
      // HTML email with plain text fallback
      const boundary = '----=_Part_' + Math.random().toString(36).substr(2, 9);
      raw = [
        `To: ${to}`,
        `Subject: ${subject}`,
        'MIME-Version: 1.0',
        `Content-Type: multipart/alternative; boundary="${boundary}"`,
        '',
        `--${boundary}`,
        'Content-Type: text/plain; charset="UTF-8"',
        'Content-Transfer-Encoding: 7bit',
        '',
        body,
        '',
        `--${boundary}`,
        'Content-Type: text/html; charset="UTF-8"',
        'Content-Transfer-Encoding: 7bit',
        '',
        html,
        '',
        `--${boundary}--`
      ].join('\r\n');
    } else {
      // Plain text email
      raw = [
        'Content-Type: text/plain; charset="UTF-8"',
        'MIME-Version: 1.0',
        `To: ${to}`,
        `Subject: ${subject}`,
        '',
        body,
      ].join('\r\n');
    }

    // Handle reply if specified
    if (reply_to_message_id) {
      const original = await gmail.users.messages.get({ userId: 'me', id: reply_to_message_id, format: 'full' });
      const headers = original.data.payload?.headers || [];
      const messageIdHeader = this.headerValue(headers, 'Message-ID');
      const threadId = original.data.threadId;
      
      if (messageIdHeader) {
        raw = raw.replace(`Subject: ${subject}`, `Subject: ${subject}\r\nIn-Reply-To: ${messageIdHeader}\r\nReferences: ${messageIdHeader}`);
      }
    }

    const encodedMessage = base64UrlEncode(raw);
    
    const draft = await gmail.users.drafts.create({
      userId: 'me',
      requestBody: {
        message: { raw: encodedMessage }
      }
    });

    return { 
      content: this.formatContent(`Draft created successfully! Draft ID: ${draft.data.id}`), 
      isText: true 
    };
  }

  async toolCreateReplyDraft(args) {
    const { message_id, body, html } = args || {};
    if (!message_id || !body) {
      throw new McpError(ErrorCode.InvalidParams, 'message_id and body are required');
    }

    const gmail = await this.getGmail();
    
    // Get the original message to extract reply information
    const original = await gmail.users.messages.get({ userId: 'me', id: message_id, format: 'full' });
    const headers = original.data.payload?.headers || [];
    const from = this.headerValue(headers, 'From');
    const subject = this.headerValue(headers, 'Subject');
    const messageIdHeader = this.headerValue(headers, 'Message-ID');
    const threadId = original.data.threadId;

    // Create proper reply subject
    const replySubject = subject?.toLowerCase().startsWith('re:') ? subject : `Re: ${subject}`;

    // Build the email content
    let raw;
    if (html) {
      // HTML email with plain text fallback
      const boundary = '----=_Part_' + Math.random().toString(36).substr(2, 9);
      raw = [
        'Content-Type: multipart/alternative; charset="UTF-8"',
        'MIME-Version: 1.0',
        `To: ${from}`,
        `Subject: ${replySubject}`,
        messageIdHeader ? `In-Reply-To: ${messageIdHeader}` : null,
        messageIdHeader ? `References: ${messageIdHeader}` : null,
        '',
        `--${boundary}`,
        'Content-Type: text/plain; charset="UTF-8"',
        'Content-Transfer-Encoding: 7bit',
        '',
        body,
        '',
        `--${boundary}`,
        'Content-Type: text/html; charset="UTF-8"',
        'Content-Transfer-Encoding: 7bit',
        '',
        html,
        '',
        `--${boundary}--`
      ].filter(Boolean).join('\r\n');
    } else {
      // Plain text email
      raw = [
        'Content-Type: text/plain; charset="UTF-8"',
        'MIME-Version: 1.0',
        `To: ${from}`,
        `Subject: ${replySubject}`,
        messageIdHeader ? `In-Reply-To: ${messageIdHeader}` : null,
        messageIdHeader ? `References: ${messageIdHeader}` : null,
        '',
        body,
      ].filter(Boolean).join('\r\n');
    }

    const encodedMessage = base64UrlEncode(raw);
    
    const draft = await gmail.users.drafts.create({
      userId: 'me',
      requestBody: {
        message: { 
          raw: encodedMessage,
          threadId: threadId
        }
      }
    });

    // Automatically get the draft to show preview
    const draftPreview = await this.toolGetDraft({ draft_id: draft.data.id });
    
    return { 
      content: this.formatContent([
        `Reply draft created successfully! Draft ID: ${draft.data.id}`,
        '',
        'Draft Preview:',
        draftPreview.content[0].text
      ]), 
      isText: true 
    };
  }

  async toolGetDraft(args) {
    const { draft_id } = args || {};
    if (!draft_id) throw new McpError(ErrorCode.InvalidParams, 'draft_id is required');

    const gmail = await this.getGmail();
    const draft = await gmail.users.drafts.get({ userId: 'me', id: draft_id, format: 'full' });
    
    const headers = draft.data.message?.payload?.headers || [];
    const to = this.headerValue(headers, 'To');
    const subject = this.headerValue(headers, 'Subject');
    const body = this.extractPlainText(draft.data.message?.payload) || '';
    const htmlBody = this.extractHtmlContent(draft.data.message?.payload) || '';

    return {
      content: this.formatContent({
        id: draft.data.id,
        to,
        subject,
        body,
        htmlBody,
        hasHtml: !!htmlBody,
        created: draft.data.message?.internalDate
      }),
      isText: true
    };
  }

  async toolListDrafts(args) {
    const max = Math.max(1, Math.min(50, Number(args?.max || 10)));
    const gmail = await this.getGmail();
    
    const list = await gmail.users.drafts.list({ userId: 'me', maxResults: max });
    const drafts = list.data.drafts || [];
    
    if (drafts.length === 0) {
      return { content: this.formatContent('No drafts found.'), isText: true };
    }

    const results = [];
    for (const draft of drafts) {
      const draftData = await gmail.users.drafts.get({ userId: 'me', id: draft.id, format: 'metadata', metadataHeaders: ['To','Subject'] });
      const headers = draftData.data.message?.payload?.headers || [];
      const to = this.headerValue(headers, 'To');
      const subject = this.headerValue(headers, 'Subject');
      
      results.push({
        id: draft.id,
        to,
        subject,
        created: draftData.data.message?.internalDate
      });
    }

    return { content: this.formatContent(results), isText: true };
  }

  async toolUpdateDraft(args) {
    const { draft_id, to, subject, body, html } = args || {};
    if (!draft_id) throw new McpError(ErrorCode.InvalidParams, 'draft_id is required');

    const gmail = await this.getGmail();
    
    // Build the updated email content
    let raw;
    if (html) {
      // HTML email with plain text fallback
      const boundary = '----=_Part_' + Math.random().toString(36).substr(2, 9);
      raw = [
        `To: ${to || 'Unknown'}`,
        `Subject: ${subject || 'No Subject'}`,
        'MIME-Version: 1.0',
        `Content-Type: multipart/alternative; boundary="${boundary}"`,
        '',
        `--${boundary}`,
        'Content-Type: text/plain; charset="UTF-8"',
        'Content-Transfer-Encoding: 7bit',
        '',
        body || '',
        '',
        `--${boundary}`,
        'Content-Type: text/html; charset="UTF-8"',
        'Content-Transfer-Encoding: 7bit',
        '',
        html,
        '',
        `--${boundary}--`
      ].join('\r\n');
    } else {
      // Plain text email
      raw = [
        'Content-Type: text/plain; charset="UTF-8"',
        'MIME-Version: 1.0',
        `To: ${to || 'Unknown'}`,
        `Subject: ${subject || 'No Subject'}`,
        '',
        body || '',
      ].join('\r\n');
    }

    const encodedMessage = base64UrlEncode(raw);
    
    await gmail.users.drafts.update({
      userId: 'me',
      id: draft_id,
      requestBody: {
        message: { raw: encodedMessage }
      }
    });

    return { 
      content: this.formatContent(`Draft ${draft_id} updated successfully!`), 
      isText: true 
    };
  }

  async toolSendDraft(args) {
    const { draft_id } = args || {};
    if (!draft_id) throw new McpError(ErrorCode.InvalidParams, 'draft_id is required');

    const gmail = await this.getGmail();
    const result = await gmail.users.drafts.send({
      userId: 'me',
      requestBody: { id: draft_id }
    });

    return { 
      content: this.formatContent(`Draft sent successfully! Message ID: ${result.data.id}`), 
      isText: true 
    };
  }

  async toolDeleteDraft(args) {
    const { draft_id } = args || {};
    if (!draft_id) throw new McpError(ErrorCode.InvalidParams, 'draft_id is required');

    const gmail = await this.getGmail();
    await gmail.users.drafts.delete({ userId: 'me', id: draft_id });

    return { 
      content: this.formatContent(`Draft ${draft_id} deleted successfully!`), 
      isText: true 
    };
  }

  parseSnoozeDate(snoozeDateStr) {
    if (!snoozeDateStr) {
      throw new McpError(ErrorCode.InvalidParams, 'snooze_date is required');
    }

    const str = String(snoozeDateStr).toLowerCase().trim();
    const now = new Date();

    // Handle ISO format: YYYY-MM-DDTHH:MM:SS
    if (str.includes('t') || str.includes('-')) {
      try {
        const date = new Date(snoozeDateStr);
        if (isNaN(date.getTime())) {
          throw new Error('Invalid date format');
        }
        return date;
      } catch (e) {
        throw new McpError(ErrorCode.InvalidParams, 'Invalid date format. Use ISO format (YYYY-MM-DDTHH:MM:SS) or relative format');
      }
    }

    // Handle Dutch date format: "23 okt 16:00" or "23 oktober 16:00"
    const dutchDateMatch = str.match(/(\d{1,2})\s+(jan|feb|mrt|apr|mei|jun|jul|aug|sep|okt|nov|dec|januari|februari|maart|april|mei|juni|juli|augustus|september|oktober|november|december)\s+(\d{1,2}):(\d{2})/);
    if (dutchDateMatch) {
      const day = parseInt(dutchDateMatch[1]);
      const monthStr = dutchDateMatch[2];
      const hours = parseInt(dutchDateMatch[3]);
      const minutes = parseInt(dutchDateMatch[4]);
      
      // Map Dutch month names to month numbers (0-based)
      const monthMap = {
        'jan': 0, 'januari': 0,
        'feb': 1, 'februari': 1,
        'mrt': 2, 'maart': 2,
        'apr': 3, 'april': 3,
        'mei': 4,
        'jun': 5, 'juni': 5,
        'jul': 6, 'juli': 6,
        'aug': 7, 'augustus': 7,
        'sep': 8, 'september': 8,
        'okt': 9, 'oktober': 9,
        'nov': 10, 'november': 10,
        'dec': 11, 'december': 11
      };
      
      const month = monthMap[monthStr];
      if (month === undefined) {
        throw new McpError(ErrorCode.InvalidParams, `Invalid Dutch month: ${monthStr}`);
      }
      
      // Create date for current year
      const year = now.getFullYear();
      const date = new Date(year, month, day, hours, minutes, 0, 0);
      
      // If the date has passed this year, assume next year
      if (date < now) {
        date.setFullYear(year + 1);
      }
      
      return date;
    }

    // Handle relative formats
    if (str.includes('tomorrow')) {
      const tomorrow = new Date(now);
      tomorrow.setDate(tomorrow.getDate() + 1);
      
      // Extract time if specified
      const timeMatch = str.match(/(\d{1,2}):?(\d{2})?\s*(am|pm)?/);
      if (timeMatch) {
        let hours = parseInt(timeMatch[1]);
        const minutes = timeMatch[2] ? parseInt(timeMatch[2]) : 0;
        const period = timeMatch[3];
        
        if (period === 'pm' && hours !== 12) hours += 12;
        if (period === 'am' && hours === 12) hours = 0;
        
        tomorrow.setHours(hours, minutes, 0, 0);
      } else {
        tomorrow.setHours(9, 0, 0, 0); // Default to 9 AM
      }
      
      return tomorrow;
    }

    if (str.includes('next monday')) {
      const nextMonday = new Date(now);
      const daysUntilMonday = (1 + 7 - now.getDay()) % 7;
      nextMonday.setDate(now.getDate() + (daysUntilMonday === 0 ? 7 : daysUntilMonday));
      nextMonday.setHours(9, 0, 0, 0);
      return nextMonday;
    }

    if (str.includes('next week')) {
      const nextWeek = new Date(now);
      nextWeek.setDate(nextWeek.getDate() + 7);
      nextWeek.setHours(9, 0, 0, 0);
      return nextWeek;
    }

    // Handle "in X hours" format
    const hoursMatch = str.match(/in\s+(\d+)\s+hours?/);
    if (hoursMatch) {
      const hours = parseInt(hoursMatch[1]);
      const future = new Date(now.getTime() + (hours * 60 * 60 * 1000));
      return future;
    }

    // Handle "in X days" format
    const daysMatch = str.match(/in\s+(\d+)\s+days?/);
    if (daysMatch) {
      const days = parseInt(daysMatch[1]);
      const future = new Date(now.getTime() + (days * 24 * 60 * 60 * 1000));
      return future;
    }

    // Handle "in X minutes" format
    const minutesMatch = str.match(/in\s+(\d+)\s+minutes?/);
    if (minutesMatch) {
      const minutes = parseInt(minutesMatch[1]);
      const future = new Date(now.getTime() + (minutes * 60 * 1000));
      return future;
    }

    throw new McpError(ErrorCode.InvalidParams, 'Unsupported date format. Use ISO format (YYYY-MM-DDTHH:MM:SS), Dutch format ("23 okt 16:00"), or relative format like "tomorrow 9am", "next monday", "in 2 hours"');
  }

  async toolSnoozeEmail(args) {
    const { id, snooze_date } = args || {};
    if (!id) throw new McpError(ErrorCode.InvalidParams, 'id is required');
    if (!snooze_date) throw new McpError(ErrorCode.InvalidParams, 'snooze_date is required');

    const gmail = await this.getGmail();
    
    // Parse the snooze date
    const snoozeDate = this.parseSnoozeDate(snooze_date);
    
    // Check if the snooze date is in the future
    const now = new Date();
    if (snoozeDate <= now) {
      throw new McpError(ErrorCode.InvalidParams, 'Snooze date must be in the future');
    }

    // Get email details for confirmation
    const msg = await gmail.users.messages.get({ userId: 'me', id, format: 'metadata', metadataHeaders: ['From','Subject','Date'] });
    const headers = msg.data.payload?.headers || [];
    const from = this.headerValue(headers, 'From');
    const subject = this.headerValue(headers, 'Subject');
    const date = this.headerValue(headers, 'Date');

    // Create a custom label for snoozed emails
    const snoozeLabelName = 'SNOOZED';
    let snoozeLabelId;
    
    try {
      // Try to find existing snooze label
      const labels = await gmail.users.labels.list({ userId: 'me' });
      const existingLabel = labels.data.labels?.find(label => label.name === snoozeLabelName);
      
      if (existingLabel) {
        snoozeLabelId = existingLabel.id;
      } else {
        // Create the snooze label
        const newLabel = await gmail.users.labels.create({
          userId: 'me',
          requestBody: {
            name: snoozeLabelName,
            labelListVisibility: 'labelShow',
            messageListVisibility: 'show'
          }
        });
        snoozeLabelId = newLabel.data.id;
      }
    } catch (error) {
      throw new McpError(ErrorCode.InternalError, `Failed to create/find snooze label: ${error.message}`);
    }

    // Add the snooze label and remove from inbox
    await gmail.users.messages.modify({
      userId: 'me',
      id,
      requestBody: {
        addLabelIds: [snoozeLabelId],
        removeLabelIds: ['INBOX']
      }
    });

    // Store snooze metadata (we'll use a simple approach with a custom header)
    // Note: Gmail API doesn't have native snooze, so we'll simulate it by archiving and adding a label
    // The user would need to manually check snoozed emails or use a separate system to "unsnooze" them
    
    const snoozeInfo = {
      snoozeDate: snoozeDate.toISOString(),
      originalInboxDate: new Date().toISOString(),
      snoozeLabel: snoozeLabelName
    };

    return {
      content: this.formatContent({
        message: 'Email snoozed successfully!',
        email: {
          subject: subject,
          from: from,
          originalDate: date
        },
        snooze: {
          until: snoozeDate.toLocaleString(),
          label: snoozeLabelName,
          note: 'Email has been archived and labeled as SNOOZED. You can search for "label:SNOOZED" to find snoozed emails.'
        }
      }),
      isText: true
    };
  }

  async toolUnsnoozeEmail(args) {
    const { id } = args || {};
    if (!id) throw new McpError(ErrorCode.InvalidParams, 'id is required');

    const gmail = await this.getGmail();
    
    // Get email details for confirmation
    const msg = await gmail.users.messages.get({ userId: 'me', id, format: 'metadata', metadataHeaders: ['From','Subject','Date'] });
    const headers = msg.data.payload?.headers || [];
    const from = this.headerValue(headers, 'From');
    const subject = this.headerValue(headers, 'Subject');
    const date = this.headerValue(headers, 'Date');

    // Find the SNOOZED label
    const labels = await gmail.users.labels.list({ userId: 'me' });
    const snoozeLabel = labels.data.labels?.find(label => label.name === 'SNOOZED');
    
    if (!snoozeLabel) {
      throw new McpError(ErrorCode.InvalidRequest, 'No SNOOZED label found. This email may not be snoozed.');
    }

    // Remove the snooze label and add back to inbox
    await gmail.users.messages.modify({
      userId: 'me',
      id,
      requestBody: {
        removeLabelIds: [snoozeLabel.id],
        addLabelIds: ['INBOX']
      }
    });

    return {
      content: this.formatContent({
        message: 'Email unsnoozed successfully!',
        email: {
          subject: subject,
          from: from,
          originalDate: date
        },
        note: 'Email has been moved back to inbox and is ready for action.'
      }),
      isText: true
    };
  }

  async toolListSnoozedEmails(args) {
    const max = Math.max(1, Math.min(50, Number(args?.max || 20)));
    const gmail = await this.getGmail();
    
    // Search for emails with SNOOZED label
    const list = await gmail.users.messages.list({ 
      userId: 'me', 
      q: 'label:SNOOZED', 
      maxResults: max 
    });
    const messages = list.data.messages || [];
    
    if (messages.length === 0) {
      return { 
        content: this.formatContent('No snoozed emails found.'), 
        isText: true 
      };
    }

    const results = [];
    for (const m of messages) {
      const msg = await gmail.users.messages.get({ 
        userId: 'me', 
        id: m.id, 
        format: 'metadata', 
        metadataHeaders: ['From','Subject','Date','Message-ID'] 
      });
      const headers = msg.data.payload?.headers || [];
      const from = this.headerValue(headers, 'From');
      const subject = this.headerValue(headers, 'Subject');
      const date = this.headerValue(headers, 'Date');
      const snippet = msg.data.snippet || '';
      
      results.push({ 
        id: m.id, 
        threadId: msg.data.threadId, 
        from, 
        subject, 
        date, 
        snippet 
      });
    }

    // Format as table similar to list_recent_unread
    const tableOutput = this.formatEmailsAsTable(results);
    return { 
      content: this.formatContent(tableOutput), 
      isText: true 
    };
  }

  // Label Management Methods
  
  async toolCreateLabel(args) {
    const { name, label_list_visibility = 'labelShow', message_list_visibility = 'show' } = args || {};
    if (!name) throw new McpError(ErrorCode.InvalidParams, 'name is required');

    const gmail = await this.getGmail();
    
    try {
      const label = await gmail.users.labels.create({
        userId: 'me',
        requestBody: {
          name,
          labelListVisibility: label_list_visibility,
          messageListVisibility: message_list_visibility
        }
      });

      return {
        content: this.formatContent({
          message: 'Label created successfully!',
          label: {
            id: label.data.id,
            name: label.data.name,
            labelListVisibility: label.data.labelListVisibility,
            messageListVisibility: label.data.messageListVisibility,
            messagesTotal: label.data.messagesTotal || 0,
            messagesUnread: label.data.messagesUnread || 0
          }
        }),
        isText: true
      };
    } catch (error) {
      if (error.message?.includes('already exists')) {
        throw new McpError(ErrorCode.InvalidRequest, `Label "${name}" already exists`);
      }
      throw new McpError(ErrorCode.InternalError, `Failed to create label: ${error.message}`);
    }
  }

  async toolCreateLabels(args) {
    const { labels } = args || {};
    if (!Array.isArray(labels) || labels.length === 0) {
      throw new McpError(ErrorCode.InvalidParams, 'labels array is required');
    }

    const gmail = await this.getGmail();
    const results = [];
    const errors = [];

    for (const labelData of labels) {
      try {
        const label = await gmail.users.labels.create({
          userId: 'me',
          requestBody: {
            name: labelData.name,
            labelListVisibility: labelData.label_list_visibility || 'labelShow',
            messageListVisibility: labelData.message_list_visibility || 'show'
          }
        });

        results.push({
          name: labelData.name,
          id: label.data.id,
          status: 'created'
        });
      } catch (error) {
        errors.push({
          name: labelData.name,
          error: error.message?.includes('already exists') ? 'Label already exists' : error.message
        });
      }
    }

    return {
      content: this.formatContent({
        message: `Batch label creation completed`,
        successful: results.length,
        failed: errors.length,
        results,
        errors: errors.length > 0 ? errors : undefined
      }),
      isText: true
    };
  }

  async toolListLabels(args) {
    const { include_system = true } = args || {};
    const gmail = await this.getGmail();
    
    const labels = await gmail.users.labels.list({ userId: 'me' });
    const allLabels = labels.data.labels || [];
    
    const filteredLabels = include_system 
      ? allLabels 
      : allLabels.filter(label => !['INBOX', 'SENT', 'DRAFT', 'SPAM', 'TRASH', 'IMPORTANT', 'STARRED', 'UNREAD'].includes(label.name));

    const formattedLabels = filteredLabels.map(label => ({
      id: label.id,
      name: label.name,
      type: label.type || 'user',
      labelListVisibility: label.labelListVisibility,
      messageListVisibility: label.messageListVisibility,
      messagesTotal: label.messagesTotal || 0,
      messagesUnread: label.messagesUnread || 0,
      threadsTotal: label.threadsTotal || 0,
      threadsUnread: label.threadsUnread || 0
    }));

    return {
      content: this.formatContent({
        message: `Found ${formattedLabels.length} labels`,
        labels: formattedLabels
      }),
      isText: true
    };
  }

  async toolGetLabel(args) {
    const { label_id } = args || {};
    if (!label_id) throw new McpError(ErrorCode.InvalidParams, 'label_id is required');

    const gmail = await this.getGmail();
    
    try {
      const label = await gmail.users.labels.get({ userId: 'me', id: label_id });
      
      return {
        content: this.formatContent({
          label: {
            id: label.data.id,
            name: label.data.name,
            type: label.data.type || 'user',
            labelListVisibility: label.data.labelListVisibility,
            messageListVisibility: label.data.messageListVisibility,
            messagesTotal: label.data.messagesTotal || 0,
            messagesUnread: label.data.messagesUnread || 0,
            threadsTotal: label.data.threadsTotal || 0,
            threadsUnread: label.data.threadsUnread || 0
          }
        }),
        isText: true
      };
    } catch (error) {
      if (error.message?.includes('not found')) {
        throw new McpError(ErrorCode.InvalidRequest, `Label with ID "${label_id}" not found`);
      }
      throw new McpError(ErrorCode.InternalError, `Failed to get label: ${error.message}`);
    }
  }

  async toolUpdateLabel(args) {
    const { label_id, name, label_list_visibility, message_list_visibility } = args || {};
    if (!label_id) throw new McpError(ErrorCode.InvalidParams, 'label_id is required');

    const gmail = await this.getGmail();
    
    const updateData = {};
    if (name) updateData.name = name;
    if (label_list_visibility) updateData.labelListVisibility = label_list_visibility;
    if (message_list_visibility) updateData.messageListVisibility = message_list_visibility;

    if (Object.keys(updateData).length === 0) {
      throw new McpError(ErrorCode.InvalidParams, 'At least one field to update is required');
    }

    try {
      const label = await gmail.users.labels.update({
        userId: 'me',
        id: label_id,
        requestBody: updateData
      });

      return {
        content: this.formatContent({
          message: 'Label updated successfully!',
          label: {
            id: label.data.id,
            name: label.data.name,
            labelListVisibility: label.data.labelListVisibility,
            messageListVisibility: label.data.messageListVisibility,
            messagesTotal: label.data.messagesTotal || 0,
            messagesUnread: label.data.messagesUnread || 0
          }
        }),
        isText: true
      };
    } catch (error) {
      if (error.message?.includes('not found')) {
        throw new McpError(ErrorCode.InvalidRequest, `Label with ID "${label_id}" not found`);
      }
      throw new McpError(ErrorCode.InternalError, `Failed to update label: ${error.message}`);
    }
  }

  async toolDeleteLabel(args) {
    const { label_id } = args || {};
    if (!label_id) throw new McpError(ErrorCode.InvalidParams, 'label_id is required');

    const gmail = await this.getGmail();
    
    try {
      await gmail.users.labels.delete({ userId: 'me', id: label_id });
      
      return {
        content: this.formatContent({
          message: `Label "${label_id}" deleted successfully!`
        }),
        isText: true
      };
    } catch (error) {
      if (error.message?.includes('not found')) {
        throw new McpError(ErrorCode.InvalidRequest, `Label with ID "${label_id}" not found`);
      }
      throw new McpError(ErrorCode.InternalError, `Failed to delete label: ${error.message}`);
    }
  }

  async toolLabelEmail(args) {
    const { email_id, label_ids } = args || {};
    if (!email_id) throw new McpError(ErrorCode.InvalidParams, 'email_id is required');
    if (!Array.isArray(label_ids) || label_ids.length === 0) {
      throw new McpError(ErrorCode.InvalidParams, 'label_ids array is required');
    }

    const gmail = await this.getGmail();
    
    try {
      await gmail.users.messages.modify({
        userId: 'me',
        id: email_id,
        requestBody: {
          addLabelIds: label_ids
        }
      });

      return {
        content: this.formatContent({
          message: `Added ${label_ids.length} label(s) to email ${email_id}`,
          email_id,
          added_labels: label_ids
        }),
        isText: true
      };
    } catch (error) {
      if (error.message?.includes('not found')) {
        throw new McpError(ErrorCode.InvalidRequest, `Email "${email_id}" not found`);
      }
      throw new McpError(ErrorCode.InternalError, `Failed to label email: ${error.message}`);
    }
  }

  async toolLabelEmails(args) {
    const { email_ids, label_ids } = args || {};
    if (!Array.isArray(email_ids) || email_ids.length === 0) {
      throw new McpError(ErrorCode.InvalidParams, 'email_ids array is required');
    }
    if (!Array.isArray(label_ids) || label_ids.length === 0) {
      throw new McpError(ErrorCode.InvalidParams, 'label_ids array is required');
    }

    const gmail = await this.getGmail();
    
    try {
      await gmail.users.messages.batchModify({
        userId: 'me',
        requestBody: {
          ids: email_ids,
          addLabelIds: label_ids
        }
      });

      return {
        content: this.formatContent({
          message: `Added ${label_ids.length} label(s) to ${email_ids.length} email(s)`,
          email_count: email_ids.length,
          added_labels: label_ids
        }),
        isText: true
      };
    } catch (error) {
      throw new McpError(ErrorCode.InternalError, `Failed to label emails: ${error.message}`);
    }
  }

  async toolUnlabelEmail(args) {
    const { email_id, label_ids } = args || {};
    if (!email_id) throw new McpError(ErrorCode.InvalidParams, 'email_id is required');
    if (!Array.isArray(label_ids) || label_ids.length === 0) {
      throw new McpError(ErrorCode.InvalidParams, 'label_ids array is required');
    }

    const gmail = await this.getGmail();
    
    try {
      await gmail.users.messages.modify({
        userId: 'me',
        id: email_id,
        requestBody: {
          removeLabelIds: label_ids
        }
      });

      return {
        content: this.formatContent({
          message: `Removed ${label_ids.length} label(s) from email ${email_id}`,
          email_id,
          removed_labels: label_ids
        }),
        isText: true
      };
    } catch (error) {
      if (error.message?.includes('not found')) {
        throw new McpError(ErrorCode.InvalidRequest, `Email "${email_id}" not found`);
      }
      throw new McpError(ErrorCode.InternalError, `Failed to unlabel email: ${error.message}`);
    }
  }

  async toolUnlabelEmails(args) {
    const { email_ids, label_ids } = args || {};
    if (!Array.isArray(email_ids) || email_ids.length === 0) {
      throw new McpError(ErrorCode.InvalidParams, 'email_ids array is required');
    }
    if (!Array.isArray(label_ids) || label_ids.length === 0) {
      throw new McpError(ErrorCode.InvalidParams, 'label_ids array is required');
    }

    const gmail = await this.getGmail();
    
    try {
      await gmail.users.messages.batchModify({
        userId: 'me',
        requestBody: {
          ids: email_ids,
          removeLabelIds: label_ids
        }
      });

      return {
        content: this.formatContent({
          message: `Removed ${label_ids.length} label(s) from ${email_ids.length} email(s)`,
          email_count: email_ids.length,
          removed_labels: label_ids
        }),
        isText: true
      };
    } catch (error) {
      throw new McpError(ErrorCode.InternalError, `Failed to unlabel emails: ${error.message}`);
    }
  }

  async toolSetEmailLabels(args) {
    const { email_id, label_ids } = args || {};
    if (!email_id) throw new McpError(ErrorCode.InvalidParams, 'email_id is required');
    if (!Array.isArray(label_ids)) {
      throw new McpError(ErrorCode.InvalidParams, 'label_ids array is required');
    }

    const gmail = await this.getGmail();
    
    try {
      // First get current labels to remove all except system labels
      const message = await gmail.users.messages.get({ userId: 'me', id: email_id, format: 'metadata' });
      const currentLabels = message.data.labelIds || [];
      
      // Keep system labels (INBOX, SENT, etc.) and remove user labels
      const systemLabels = ['INBOX', 'SENT', 'DRAFT', 'SPAM', 'TRASH', 'IMPORTANT', 'STARRED', 'UNREAD'];
      const userLabelsToRemove = currentLabels.filter(labelId => !systemLabels.includes(labelId));
      
      await gmail.users.messages.modify({
        userId: 'me',
        id: email_id,
        requestBody: {
          removeLabelIds: userLabelsToRemove,
          addLabelIds: label_ids
        }
      });

      return {
        content: this.formatContent({
          message: `Set labels for email ${email_id}`,
          email_id,
          new_labels: label_ids,
          removed_labels: userLabelsToRemove
        }),
        isText: true
      };
    } catch (error) {
      if (error.message?.includes('not found')) {
        throw new McpError(ErrorCode.InvalidRequest, `Email "${email_id}" not found`);
      }
      throw new McpError(ErrorCode.InternalError, `Failed to set email labels: ${error.message}`);
    }
  }

  async toolGetEmailLabels(args) {
    const { email_id } = args || {};
    if (!email_id) throw new McpError(ErrorCode.InvalidParams, 'email_id is required');

    const gmail = await this.getGmail();
    
    try {
      const message = await gmail.users.messages.get({ userId: 'me', id: email_id, format: 'metadata' });
      const labelIds = message.data.labelIds || [];
      
      // Get label details
      const labels = await gmail.users.labels.list({ userId: 'me' });
      const allLabels = labels.data.labels || [];
      
      const emailLabels = labelIds.map(labelId => {
        const label = allLabels.find(l => l.id === labelId);
        return {
          id: labelId,
          name: label?.name || 'Unknown',
          type: label?.type || 'unknown'
        };
      });

      return {
        content: this.formatContent({
          email_id,
          labels: emailLabels,
          label_count: emailLabels.length
        }),
        isText: true
      };
    } catch (error) {
      if (error.message?.includes('not found')) {
        throw new McpError(ErrorCode.InvalidRequest, `Email "${email_id}" not found`);
      }
      throw new McpError(ErrorCode.InternalError, `Failed to get email labels: ${error.message}`);
    }
  }

  async toolSearchByLabel(args) {
    const { label_names, label_ids, additional_query = '', max = 50 } = args || {};
    
    if (!label_names && !label_ids) {
      throw new McpError(ErrorCode.InvalidParams, 'Either label_names or label_ids is required');
    }

    const gmail = await this.getGmail();
    
    // Build search query
    let query = '';
    
    if (label_names && Array.isArray(label_names)) {
      const labelQueries = label_names.map(name => `label:"${name}"`);
      query += labelQueries.join(' ');
    }
    
    if (label_ids && Array.isArray(label_ids)) {
      const labelIdQueries = label_ids.map(id => `label:${id}`);
      query += (query ? ' ' : '') + labelIdQueries.join(' ');
    }
    
    if (additional_query) {
      query += (query ? ' ' : '') + additional_query;
    }

    const maxResults = Math.max(1, Math.min(100, Number(max)));
    
    try {
      const list = await gmail.users.messages.list({ 
        userId: 'me', 
        q: query, 
        maxResults 
      });
      const messages = list.data.messages || [];
      
      if (messages.length === 0) {
        return { 
          content: this.formatContent(`No emails found for label search: "${query}"`), 
          isText: true 
        };
      }

      const results = [];
      for (const m of messages) {
        const msg = await gmail.users.messages.get({ 
          userId: 'me', 
          id: m.id, 
          format: 'metadata', 
          metadataHeaders: ['From','Subject','Date','Message-ID'] 
        });
        const headers = msg.data.payload?.headers || [];
        const from = this.headerValue(headers, 'From');
        const subject = this.headerValue(headers, 'Subject');
        const date = this.headerValue(headers, 'Date');
        const snippet = msg.data.snippet || '';
        const labels = msg.data.labelIds || [];
        const unread = labels.includes('UNREAD');
        
        results.push({ 
          id: m.id, 
          threadId: msg.data.threadId, 
          from, 
          subject, 
          date, 
          snippet,
          unread,
          labels: labels
        });
      }

      return { 
        content: this.formatContent({
          query: query,
          count: results.length,
          emails: results
        }), 
        isText: true 
      };
    } catch (error) {
      throw new McpError(ErrorCode.InternalError, `Failed to search by label: ${error.message}`);
    }
  }

  async toolBulkLabelEmails(args) {
    const { query, label_ids, operation = 'add', batch_size = 100, dry_run = false } = args || {};
    if (!query) throw new McpError(ErrorCode.InvalidParams, 'query is required');
    if (!Array.isArray(label_ids) || label_ids.length === 0) {
      throw new McpError(ErrorCode.InvalidParams, 'label_ids array is required');
    }

    const batchSize = Math.max(10, Math.min(500, Number(batch_size)));
    const gmail = await this.getGmail();
    
    // First, get total count
    const list = await gmail.users.messages.list({ userId: 'me', q: query, maxResults: 1 });
    const totalCount = list.data.resultSizeEstimate || 0;
    
    if (totalCount === 0) {
      return { 
        content: this.formatContent(`No emails found matching query: "${query}"`), 
        isText: true 
      };
    }

    if (dry_run) {
      return { 
        content: this.formatContent({
          action: 'DRY RUN - No emails were labeled',
          query: query,
          operation: operation,
          total_estimated: totalCount,
          batch_size: batchSize,
          estimated_batches: Math.ceil(totalCount / batchSize),
          labels_to_apply: label_ids,
          note: 'This is an estimate. Actual count may vary.'
        }), 
        isText: true 
      };
    }

    // Process in batches
    let processedCount = 0;
    let labeledCount = 0;
    let errorCount = 0;
    const errors = [];
    let nextPageToken = null;

    while (processedCount < totalCount) {
      const listParams = { 
        userId: 'me', 
        q: query, 
        maxResults: batchSize 
      };
      if (nextPageToken) {
        listParams.pageToken = nextPageToken;
      }
      
      const batchList = await gmail.users.messages.list(listParams);
      const messages = batchList.data.messages || [];
      nextPageToken = batchList.data.nextPageToken;
      
      if (messages.length === 0) break;
      
      const messageIds = messages.map(m => m.id);
      
      try {
        const modifyRequest = {
          userId: 'me',
          requestBody: { ids: messageIds }
        };

        if (operation === 'add') {
          modifyRequest.requestBody.addLabelIds = label_ids;
        } else if (operation === 'remove') {
          modifyRequest.requestBody.removeLabelIds = label_ids;
        } else if (operation === 'replace') {
          // For replace, we need to get current labels first and remove user labels
          const systemLabels = ['INBOX', 'SENT', 'DRAFT', 'SPAM', 'TRASH', 'IMPORTANT', 'STARRED', 'UNREAD'];
          const messages = await Promise.all(
            messageIds.map(id => gmail.users.messages.get({ userId: 'me', id, format: 'metadata' }))
          );
          
          const allCurrentLabels = messages.map(msg => msg.data.labelIds || []);
          const userLabelsToRemove = allCurrentLabels.flat().filter(labelId => !systemLabels.includes(labelId));
          
          modifyRequest.requestBody.removeLabelIds = [...new Set(userLabelsToRemove)];
          modifyRequest.requestBody.addLabelIds = label_ids;
        }

        await gmail.users.messages.batchModify(modifyRequest);
        labeledCount += messageIds.length;
      } catch (error) {
        errorCount += messageIds.length;
        errors.push({ 
          batch: Math.floor(processedCount / batchSize) + 1, 
          error: error.message,
          count: messageIds.length
        });
      }
      
      processedCount += messages.length;
      
      if (!nextPageToken) break;
    }

    return { 
      content: this.formatContent({
        action: `Bulk labeling operation completed`,
        query: query,
        operation: operation,
        total_processed: processedCount,
        successful: labeledCount,
        failed: errorCount,
        batch_size: batchSize,
        batches_processed: Math.ceil(processedCount / batchSize),
        labels_applied: label_ids,
        errors: errors.length > 0 ? errors : undefined,
        performance_note: 'Used batch operations for optimal speed'
      }), 
      isText: true 
    };
  }

  async toolAutoLabelEmails(args) {
    const { rules, dry_run = false } = args || {};
    if (!Array.isArray(rules) || rules.length === 0) {
      throw new McpError(ErrorCode.InvalidParams, 'rules array is required');
    }

    const gmail = await this.getGmail();
    const results = [];
    const errors = [];

    for (const rule of rules) {
      try {
        // Build search query for this rule
        let query = '';
        if (rule.query) {
          query = rule.query;
        } else {
          const conditions = [];
          if (rule.sender_pattern) {
            conditions.push(`from:${rule.sender_pattern}`);
          }
          if (rule.subject_pattern) {
            conditions.push(`subject:"${rule.subject_pattern}"`);
          }
          query = conditions.join(' ');
        }

        if (!query) {
          errors.push({
            rule: rule.label_name,
            error: 'No search criteria provided (need sender_pattern, subject_pattern, or query)'
          });
          continue;
        }

        // Find or create the label
        let labelId;
        const labels = await gmail.users.labels.list({ userId: 'me' });
        const existingLabel = labels.data.labels?.find(label => label.name === rule.label_name);
        
        if (existingLabel) {
          labelId = existingLabel.id;
        } else {
          const newLabel = await gmail.users.labels.create({
            userId: 'me',
            requestBody: {
              name: rule.label_name,
              labelListVisibility: 'labelShow',
              messageListVisibility: 'show'
            }
          });
          labelId = newLabel.data.id;
        }

        // Search for emails matching this rule
        const list = await gmail.users.messages.list({ 
          userId: 'me', 
          q: query, 
          maxResults: 100 
        });
        const messages = list.data.messages || [];

        if (messages.length === 0) {
          results.push({
            rule: rule.label_name,
            label_id: labelId,
            emails_found: 0,
            status: 'no_emails_found'
          });
          continue;
        }

        if (dry_run) {
          results.push({
            rule: rule.label_name,
            label_id: labelId,
            emails_found: messages.length,
            status: 'dry_run',
            query: query
          });
        } else {
          // Apply the label
          const messageIds = messages.map(m => m.id);
          await gmail.users.messages.batchModify({
            userId: 'me',
            requestBody: {
              ids: messageIds,
              addLabelIds: [labelId]
            }
          });

          results.push({
            rule: rule.label_name,
            label_id: labelId,
            emails_found: messages.length,
            emails_labeled: messageIds.length,
            status: 'labeled',
            query: query
          });
        }
      } catch (error) {
        errors.push({
          rule: rule.label_name,
          error: error.message
        });
      }
    }

    return {
      content: this.formatContent({
        message: dry_run ? 'Auto-labeling dry run completed' : 'Auto-labeling completed',
        rules_processed: rules.length,
        successful: results.length,
        failed: errors.length,
        results,
        errors: errors.length > 0 ? errors : undefined
      }),
      isText: true
    };
  }

  async toolLabelEmailsByQuery(args) {
    const { query, label_name, label_id, max = 100, dry_run = false } = args || {};
    if (!query) throw new McpError(ErrorCode.InvalidParams, 'query is required');
    if (!label_name && !label_id) {
      throw new McpError(ErrorCode.InvalidParams, 'Either label_name or label_id is required');
    }

    const gmail = await this.getGmail();
    
    // Find or create the label
    let targetLabelId = label_id;
    if (label_name) {
      const labels = await gmail.users.labels.list({ userId: 'me' });
      const existingLabel = labels.data.labels?.find(label => label.name === label_name);
      
      if (existingLabel) {
        targetLabelId = existingLabel.id;
      } else {
        const newLabel = await gmail.users.labels.create({
          userId: 'me',
          requestBody: {
            name: label_name,
            labelListVisibility: 'labelShow',
            messageListVisibility: 'show'
          }
        });
        targetLabelId = newLabel.data.id;
      }
    }

    const maxResults = Math.max(1, Math.min(500, Number(max)));
    
    // Search for emails
    const list = await gmail.users.messages.list({ 
      userId: 'me', 
      q: query, 
      maxResults 
    });
    const messages = list.data.messages || [];
    
    if (messages.length === 0) {
      return { 
        content: this.formatContent(`No emails found matching query: "${query}"`), 
        isText: true 
      };
    }

    if (dry_run) {
      return { 
        content: this.formatContent({
          action: 'DRY RUN - No emails were labeled',
          query: query,
          label_name: label_name,
          label_id: targetLabelId,
          emails_found: messages.length,
          note: 'This shows what would be labeled without actually labeling'
        }), 
        isText: true 
      };
    }

    // Apply the label
    const messageIds = messages.map(m => m.id);
    await gmail.users.messages.batchModify({
      userId: 'me',
      requestBody: {
        ids: messageIds,
        addLabelIds: [targetLabelId]
      }
    });

    return {
      content: this.formatContent({
        message: 'Emails labeled successfully!',
        query: query,
        label_name: label_name,
        label_id: targetLabelId,
        emails_labeled: messageIds.length
      }),
      isText: true
    };
  }

  async toolGetLabelStatistics(args) {
    const { label_ids } = args || {};
    const gmail = await this.getGmail();
    
    const labels = await gmail.users.labels.list({ userId: 'me' });
    const allLabels = labels.data.labels || [];
    
    const targetLabels = label_ids 
      ? allLabels.filter(label => label_ids.includes(label.id))
      : allLabels.filter(label => label.type === 'user'); // Only user labels by default

    const statistics = targetLabels.map(label => ({
      id: label.id,
      name: label.name,
      type: label.type || 'user',
      messages_total: label.messagesTotal || 0,
      messages_unread: label.messagesUnread || 0,
      threads_total: label.threadsTotal || 0,
      threads_unread: label.threadsUnread || 0,
      read_percentage: label.messagesTotal > 0 
        ? Math.round(((label.messagesTotal - (label.messagesUnread || 0)) / label.messagesTotal) * 100)
        : 0
    }));

    const totalStats = {
      total_labels: statistics.length,
      total_messages: statistics.reduce((sum, stat) => sum + stat.messages_total, 0),
      total_unread: statistics.reduce((sum, stat) => sum + stat.messages_unread, 0),
      total_threads: statistics.reduce((sum, stat) => sum + stat.threads_total, 0),
      overall_read_percentage: statistics.length > 0 
        ? Math.round(statistics.reduce((sum, stat) => sum + stat.read_percentage, 0) / statistics.length)
        : 0
    };

    return {
      content: this.formatContent({
        message: 'Label statistics retrieved',
        statistics,
        totals: totalStats
      }),
      isText: true
    };
  }

  async toolCleanupUnusedLabels(args) {
    const { dry_run = true } = args || {};
    const gmail = await this.getGmail();
    
    const labels = await gmail.users.labels.list({ userId: 'me' });
    const allLabels = labels.data.labels || [];
    
    // Filter to user labels only (exclude system labels)
    const userLabels = allLabels.filter(label => 
      label.type === 'user' && 
      !['INBOX', 'SENT', 'DRAFT', 'SPAM', 'TRASH', 'IMPORTANT', 'STARRED', 'UNREAD'].includes(label.name)
    );

    const unusedLabels = userLabels.filter(label => 
      (label.messagesTotal || 0) === 0 && (label.threadsTotal || 0) === 0
    );

    if (unusedLabels.length === 0) {
      return {
        content: this.formatContent({
          message: 'No unused labels found',
          total_user_labels: userLabels.length,
          unused_labels: 0
        }),
        isText: true
      };
    }

    if (dry_run) {
      return {
        content: this.formatContent({
          action: 'DRY RUN - No labels were deleted',
          total_user_labels: userLabels.length,
          unused_labels_found: unusedLabels.length,
          unused_labels: unusedLabels.map(label => ({
            id: label.id,
            name: label.name,
            messages_total: label.messagesTotal || 0,
            threads_total: label.threadsTotal || 0
          })),
          note: 'Set dry_run=false to actually delete these labels'
        }),
        isText: true
      };
    }

    // Actually delete unused labels
    const deletedLabels = [];
    const errors = [];

    for (const label of unusedLabels) {
      try {
        await gmail.users.labels.delete({ userId: 'me', id: label.id });
        deletedLabels.push({
          id: label.id,
          name: label.name
        });
      } catch (error) {
        errors.push({
          id: label.id,
          name: label.name,
          error: error.message
        });
      }
    }

    return {
      content: this.formatContent({
        message: 'Label cleanup completed',
        total_user_labels: userLabels.length,
        deleted_labels: deletedLabels.length,
        failed_deletions: errors.length,
        deleted_labels: deletedLabels,
        errors: errors.length > 0 ? errors : undefined
      }),
      isText: true
    };
  }

  // Auto-Labeling Rules Management Methods

  async toolListAutoLabelingRules(args) {
    const rules = await this.loadAutoLabelingRules();
    
    const formattedRules = rules.map((rule, index) => ({
      index: index,
      label_name: rule.label_name,
      sender_pattern: rule.sender_pattern || null,
      subject_pattern: rule.subject_pattern || null,
      query: rule.query || null,
      enabled: rule.enabled !== false, // Default to true
      created: rule.created || 'Unknown',
      last_run: rule.last_run || 'Never'
    }));

    return {
      content: this.formatContent({
        message: `Found ${formattedRules.length} auto-labeling rules`,
        config_file: this.getRulesConfigPath(),
        rules: formattedRules
      }),
      isText: true
    };
  }

  async toolAddAutoLabelingRule(args) {
    const { label_name, sender_pattern, subject_pattern, subject_contains, query, enabled = true } = args || {};
    if (!label_name) throw new McpError(ErrorCode.InvalidParams, 'label_name is required');

    const rules = await this.loadAutoLabelingRules();
    
    // Build the search query for this rule
    let searchQuery = '';
    if (query) {
      searchQuery = query;
    } else {
      const conditions = [];
      if (sender_pattern) {
        conditions.push(`from:${sender_pattern}`);
      }
      if (subject_pattern) {
        conditions.push(`subject:"${subject_pattern}"`);
      }
      if (subject_contains && Array.isArray(subject_contains) && subject_contains.length > 0) {
        const subjectConditions = subject_contains.map(text => `subject:"${text}"`);
        conditions.push(`(${subjectConditions.join(' OR ')})`);
      }
      searchQuery = conditions.join(' ');
    }

    if (!searchQuery) {
      throw new McpError(ErrorCode.InvalidParams, 'No search criteria provided (need sender_pattern, subject_pattern, subject_contains, or query)');
    }

    const newRule = {
      label_name,
      sender_pattern: sender_pattern || null,
      subject_pattern: subject_pattern || null,
      subject_contains: subject_contains || null,
      query: searchQuery,
      enabled: enabled,
      created: new Date().toISOString(),
      last_run: null
    };

    rules.push(newRule);
    await this.saveAutoLabelingRules(rules);

    return {
      content: this.formatContent({
        message: 'Auto-labeling rule added successfully!',
        rule: {
          index: rules.length - 1,
          ...newRule
        },
        total_rules: rules.length
      }),
      isText: true
    };
  }

  async toolRemoveAutoLabelingRule(args) {
    const { rule_index } = args || {};
    if (typeof rule_index !== 'number' || rule_index < 0) {
      throw new McpError(ErrorCode.InvalidParams, 'rule_index must be a non-negative number');
    }

    const rules = await this.loadAutoLabelingRules();
    
    if (rule_index >= rules.length) {
      throw new McpError(ErrorCode.InvalidRequest, `Rule index ${rule_index} is out of range. Available rules: 0-${rules.length - 1}`);
    }

    const removedRule = rules.splice(rule_index, 1)[0];
    await this.saveAutoLabelingRules(rules);

    return {
      content: this.formatContent({
        message: 'Auto-labeling rule removed successfully!',
        removed_rule: {
          index: rule_index,
          ...removedRule
        },
        remaining_rules: rules.length
      }),
      isText: true
    };
  }

  async toolUpdateAutoLabelingRule(args) {
    const { rule_index, label_name, sender_pattern, subject_pattern, subject_contains, query, enabled } = args || {};
    if (typeof rule_index !== 'number' || rule_index < 0) {
      throw new McpError(ErrorCode.InvalidParams, 'rule_index must be a non-negative number');
    }

    const rules = await this.loadAutoLabelingRules();
    
    if (rule_index >= rules.length) {
      throw new McpError(ErrorCode.InvalidRequest, `Rule index ${rule_index} is out of range. Available rules: 0-${rules.length - 1}`);
    }

    const rule = rules[rule_index];
    const originalRule = { ...rule };

    // Update fields if provided
    if (label_name !== undefined) rule.label_name = label_name;
    if (sender_pattern !== undefined) rule.sender_pattern = sender_pattern;
    if (subject_pattern !== undefined) rule.subject_pattern = subject_pattern;
    if (subject_contains !== undefined) rule.subject_contains = subject_contains;
    if (query !== undefined) rule.query = query;
    if (enabled !== undefined) rule.enabled = enabled;

    // Rebuild search query if patterns changed
    if (sender_pattern !== undefined || subject_pattern !== undefined || subject_contains !== undefined) {
      if (rule.query && !query) {
        // Only rebuild if no explicit query was provided
        const conditions = [];
        if (rule.sender_pattern) {
          conditions.push(`from:${rule.sender_pattern}`);
        }
        if (rule.subject_pattern) {
          conditions.push(`subject:"${rule.subject_pattern}"`);
        }
        if (rule.subject_contains && Array.isArray(rule.subject_contains) && rule.subject_contains.length > 0) {
          const subjectConditions = rule.subject_contains.map(text => `subject:"${text}"`);
          conditions.push(`(${subjectConditions.join(' OR ')})`);
        }
        rule.query = conditions.join(' ');
      }
    }

    rule.updated = new Date().toISOString();
    await this.saveAutoLabelingRules(rules);

    return {
      content: this.formatContent({
        message: 'Auto-labeling rule updated successfully!',
        rule_index: rule_index,
        changes: {
          original: originalRule,
          updated: rule
        }
      }),
      isText: true
    };
  }

  async toolRunAutoLabelingRules(args) {
    const { dry_run = false, max_per_rule = 100, batch_size = 100, max_batches = 10 } = args || {};
    const rules = await this.loadAutoLabelingRules();
    
    const enabledRules = rules.filter(rule => rule.enabled !== false);
    
    if (enabledRules.length === 0) {
      return {
        content: this.formatContent({
          message: 'No enabled auto-labeling rules found',
          total_rules: rules.length,
          enabled_rules: 0
        }),
        isText: true
      };
    }

    const gmail = await this.getGmail();
    const results = [];
    const errors = [];

    for (const rule of enabledRules) {
      try {
        // Find or create the label
        let labelId;
        const labels = await gmail.users.labels.list({ userId: 'me' });
        const existingLabel = labels.data.labels?.find(label => label.name === rule.label_name);
        
        if (existingLabel) {
          labelId = existingLabel.id;
        } else {
          const newLabel = await gmail.users.labels.create({
            userId: 'me',
            requestBody: {
              name: rule.label_name,
              labelListVisibility: 'labelShow',
              messageListVisibility: 'show'
            }
          });
          labelId = newLabel.data.id;
        }

        // Generate query from subject_contains if query is null
        let searchQuery = rule.query;
        if (!searchQuery && rule.subject_contains && Array.isArray(rule.subject_contains) && rule.subject_contains.length > 0) {
          const subjectConditions = rule.subject_contains.map(text => `subject:"${text}"`);
          searchQuery = `(${subjectConditions.join(' OR ')})`;
        }

        if (!searchQuery) {
          results.push({
            rule: rule.label_name,
            label_id: labelId,
            emails_found: 0,
            status: 'no_search_criteria'
          });
          continue;
        }

        // Process emails in batches
        let totalEmailsFound = 0;
        let totalEmailsLabeled = 0;
        let nextPageToken = null;
        let batchCount = 0;
        const allMessageIds = [];

        while (batchCount < max_batches) {
          // Search for emails matching this rule
          const listParams = { 
            userId: 'me', 
            q: searchQuery, 
            maxResults: Math.min(batch_size, 500)
          };
          
          if (nextPageToken) {
            listParams.pageToken = nextPageToken;
          }

          const list = await gmail.users.messages.list(listParams);
          const messages = list.data.messages || [];
          
          if (messages.length === 0) {
            break; // No more emails
          }

          totalEmailsFound += messages.length;
          allMessageIds.push(...messages.map(m => m.id));
          
          // Check if there are more pages
          nextPageToken = list.data.nextPageToken;
          if (!nextPageToken) {
            break; // No more pages
          }
          
          batchCount++;
        }

        if (totalEmailsFound === 0) {
          results.push({
            rule: rule.label_name,
            label_id: labelId,
            emails_found: 0,
            status: 'no_emails_found'
          });
          continue;
        }

        if (dry_run) {
          results.push({
            rule: rule.label_name,
            label_id: labelId,
            emails_found: totalEmailsFound,
            batches_processed: batchCount + 1,
            status: 'dry_run',
            query: searchQuery
          });
        } else {
          // Apply labels in batches (Gmail API limit is 1000 IDs per batch)
          const maxBatchSize = 1000;
          for (let i = 0; i < allMessageIds.length; i += maxBatchSize) {
            const batchIds = allMessageIds.slice(i, i + maxBatchSize);
            await gmail.users.messages.batchModify({
              userId: 'me',
              requestBody: {
                ids: batchIds,
                addLabelIds: [labelId]
              }
            });
            totalEmailsLabeled += batchIds.length;
          }

          // Update last_run timestamp
          rule.last_run = new Date().toISOString();
          await this.saveAutoLabelingRules(rules);

          results.push({
            rule: rule.label_name,
            label_id: labelId,
            emails_found: totalEmailsFound,
            emails_labeled: totalEmailsLabeled,
            batches_processed: batchCount + 1,
            status: 'labeled',
            query: searchQuery
          });
        }
      } catch (error) {
        errors.push({
          rule: rule.label_name,
          error: error.message
        });
      }
    }

    return {
      content: this.formatContent({
        message: dry_run ? 'Auto-labeling rules dry run completed' : 'Auto-labeling rules executed successfully',
        total_rules: rules.length,
        enabled_rules: enabledRules.length,
        processed: results.length,
        failed: errors.length,
        results,
        errors: errors.length > 0 ? errors : undefined
      }),
      isText: true
    };
  }

  async toolExportAutoLabelingRules(args) {
    const { file_path } = args || {};
    const rules = await this.loadAutoLabelingRules();
    
    const exportData = {
      version: '1.0.0',
      exported: new Date().toISOString(),
      rules: rules
    };

    const exportPath = file_path || path.join(this.getConfigDir(), `auto-labeling-rules-backup-${new Date().toISOString().split('T')[0]}.json`);
    
    await fs.writeFile(exportPath, JSON.stringify(exportData, null, 2), 'utf8');

    return {
      content: this.formatContent({
        message: 'Auto-labeling rules exported successfully!',
        export_path: exportPath,
        rules_exported: rules.length,
        file_size: `${Math.round(JSON.stringify(exportData).length / 1024)} KB`
      }),
      isText: true
    };
  }

  async toolImportAutoLabelingRules(args) {
    const { file_path, merge = false } = args || {};
    if (!file_path) throw new McpError(ErrorCode.InvalidParams, 'file_path is required');

    try {
      const raw = await fs.readFile(file_path, 'utf8');
      const importData = JSON.parse(raw);
      
      if (!importData.rules || !Array.isArray(importData.rules)) {
        throw new McpError(ErrorCode.InvalidRequest, 'Invalid import file format. Expected "rules" array.');
      }

      let rules;
      if (merge) {
        const existingRules = await this.loadAutoLabelingRules();
        rules = [...existingRules, ...importData.rules];
      } else {
        rules = importData.rules;
      }

      await this.saveAutoLabelingRules(rules);

      return {
        content: this.formatContent({
          message: 'Auto-labeling rules imported successfully!',
          import_path: file_path,
          rules_imported: importData.rules.length,
          total_rules: rules.length,
          merge_mode: merge,
          imported_version: importData.version || 'Unknown'
        }),
        isText: true
      };
    } catch (error) {
      if (error instanceof McpError) throw error;
      throw new McpError(ErrorCode.InternalError, `Failed to import rules: ${error.message}`);
    }
  }

  async run() {
    const transport = new StdioServerTransport();
    await this.server.connect(transport);
    console.error('Gmail MCP server running on stdio');
  }
}

const server = new GmailMcpServer();
server.run().catch(console.error);


