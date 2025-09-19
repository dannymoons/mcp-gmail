# Gmail MCP Server

A comprehensive Gmail integration for Cursor IDE that provides email management, collaborative drafting, and intelligent email crafting capabilities.

## Features

- **Email Management**: Read, reply, archive, delete emails
- **Collaborative Drafting**: Create, review, and edit drafts before sending
- **Smart Email Crafting**: AI-powered email composition with your personal style
- **Bulk Operations**: Efficiently manage multiple emails
- **Web UI**: Clean, modern interface for email management
- **Search & Filter**: Advanced Gmail search capabilities

## Setup

### 1. Google OAuth Credentials
- Create Google OAuth Desktop credentials in Google Cloud Console
- Add redirect URI: `http://127.0.0.1:53682/oauth2callback`
- Provide credentials via environment variables or config file

### 2. Cursor Configuration
Register in `~/.cursor/mcp.json`:
```json
{
  "mcpServers": {
    "gmail": {
      "command": "node",
      "args": ["/Users/danny/development/mcp-servers/mcp-gmail/index.js"],
      "env": {
        "GOOGLE_CLIENT_ID": "<your_client_id>",
        "GOOGLE_CLIENT_SECRET": "<your_client_secret>"
      }
    }
  }
}
```

## Configuration

### Auto-Labeling Rules
Auto-labeling rules are stored in `auto-labeling-rules.json` within the project directory. This file contains:
- Rule definitions with search criteria
- Enable/disable status for each rule
- Creation and last-run timestamps
- Version control metadata

## Tools

### Authentication
- `start_oauth({ port? })` - Start OAuth flow, returns authorization URL
- `auth_status()` - Check current authentication status

### Email Reading
- `list_unread({ max? })` - List recent unread emails
- `list_recent_unread({ days?, max? })` - List recent unread emails in table format (default: 7 days)
- `get_message({ id })` - Get full message content
- `search_emails({ query, max? })` - Search emails with Gmail syntax

#### list_recent_unread Method
This new method provides a clean table format for viewing recent unread emails:

**Parameters:**
- `days` (string, optional): Number of days to look back. Supports multiple formats:
  - `"2"` - 2 days
  - `"3d"` - 3 days  
  - `"2 days"` - 2 days
  - `"--2"` - 2 days
  - Default: `"7"` (7 days)
- `max` (number, optional): Maximum emails to return (default: 50, max: 100)

**Output Format:**
```
| ID | Date | Sender | Subject |
|---|---|---|---|
| 1 | 18/09 16:56 | Dumith Salinda (GitHub) | Re: [pixeltoplanet/carbonfooter-plugin] feat(cache) |
| 2 | 18/09 21:40 | Ivan at Notion | Notion 3.0: Agents |

**Total: 2 unread emails**
```

**Examples:**
- `list_recent_unread()` - Last 7 days, max 50 emails
- `list_recent_unread({days: "2"})` - Last 2 days
- `list_recent_unread({days: "3d", max: 20})` - Last 3 days, max 20 emails

### Email Management
- `mark_as_read({ id })` - Mark message as read
- `batch_archive({ ids[] })` - Archive multiple messages
- `batch_delete({ ids[], permanent? })` - Delete multiple messages
- `delete_email({ id, permanent? })` - Delete single email
- `delete_emails_by_query({ query, max?, permanent?, dry_run? })` - Delete emails by search query
- `bulk_delete_emails({ query, batch_size?, permanent?, dry_run? })` - High-performance bulk delete

### Email Composition
- `reply_to_message({ message_id, body })` - Send immediate reply
- `create_draft({ to, subject, body, html?, reply_to_message_id? })` - Create draft email
- `get_draft({ draft_id })` - Review draft content
- `list_drafts({ max? })` - List all drafts
- `update_draft({ draft_id, to?, subject?, body?, html? })` - Update existing draft
- `send_draft({ draft_id })` - Send draft email
- `delete_draft({ draft_id })` - Delete draft

### User Interface
- `start_ui({ port?, query?, max? })` - Launch web UI at http://127.0.0.1:53750/
- `stop_ui()` - Stop the web UI server

### Auto-Labeling Rules Management
- `list_auto_labeling_rules()` - List all stored auto-labeling rules
- `add_auto_labeling_rule({ label_name, sender_pattern?, subject_pattern?, query?, enabled? })` - Add new auto-labeling rule
- `remove_auto_labeling_rule({ rule_index })` - Remove rule by index
- `update_auto_labeling_rule({ rule_index, label_name?, sender_pattern?, subject_pattern?, query?, enabled? })` - Update existing rule
- `run_auto_labeling_rules({ dry_run?, max_per_rule? })` - Execute all enabled rules
- `export_auto_labeling_rules({ file_path? })` - Export rules to JSON file
- `import_auto_labeling_rules({ file_path, merge? })` - Import rules from JSON file

## Usage

### Basic Workflow
1. Call `start_oauth` and authorize in browser
2. Use `list_unread` or `search_emails` to find emails
3. Use `get_message` to read full content
4. Use `reply_to_message` for quick replies or `create_draft` for collaborative writing

### Collaborative Email Crafting
1. `create_draft` - Start with basic content
2. `get_draft` - Review the draft
3. `update_draft` - Refine the content together
4. `send_draft` - Send when perfect

### Web Interface
- Call `start_ui` to open the modern web interface
- Browse emails, select multiple items, perform bulk actions
- Rich text editor for composing replies

## Danny's Email Writing Style & Templates

*Based on analysis of 30 recent sent emails*

### Writing Style Preferences
- **Tone**: Conversational yet professional, direct and honest
- **Language**: Dutch (80% of emails) for personal/business relationships, English for international/technical contexts
- **Structure**: Clear problem-solution format, bullet points for achievements, short paragraphs
- **Sign-off**: "Met vriendelijke groet, Danny Moons" (standard)
- **Signature**: Always includes email (danny@moonsio.nl) and phone (0628509910)

### Actual Email Patterns

#### Opening Styles
- **Dutch Personal**: "Hoi [Name]," (most common)
- **English**: "Hi [Name]," (international/technical)
- **Casual**: "Hey [Name]," (close business relationships)

#### Content Structure
- **Problem Identification**: Clear explanation of issues
- **Solution-Oriented**: Always provides actionable next steps
- **Educational Approach**: Explains what went wrong and how to prevent it
- **Collaborative Tone**: "Laten we samen bepalen wat prioriteit heeft"
- **Timeline Management**: Clear expectations ("na mijn vakantie", "begin oktober")

#### Closing Patterns
- **Standard**: "Met vriendelijke groet, Danny Moons"
- **Contact Info**: Always includes email and phone
- **Personal Touch**: "Ik hoor het graag" / "Ik hoor graag wat je ervan vind"
- **Confirmation Requests**: "Kun je controleren of het zo goed staat?"

### Email Templates

#### Dutch Personal Template (Based on Actual Usage)
```
Hoi {name},

{personal greeting and context}

{main content with bullet points for achievements}

Wat we hebben bereikt:
• {achievement 1}
• {achievement 2}
• {achievement 3}

Ik hoor graag wat je ervan vind.

Met vriendelijke groet,
Danny Moons
E: danny@moonsio.nl
T: 0628509910
```

#### Dutch Business Template (Based on Actual Usage)
```
Hoi {name},

{problem identification and context}

{clear solution or next steps}

{timeline or expectations}

Ik hoor het graag. Voor nu alvast een fijn weekend en ik spreek je ongetwijfeld na mijn korte vakantie.

Met vriendelijke groet,
Danny Moons
E: danny@moonsio.nl
T: 0628509910
```

#### Technical Support Template
```
Hi {name},

Waarschijnlijk heb je {issue} per ongeluk {cause}, waardoor {symptom}. Ik heb het nu {solution} door {action}. Kun je controleren of het zo goed staat?

Wil je dat ik nog andere {items} aanpas of is dit voldoende voor nu?

Met vriendelijke groet,
Danny Moons
E: danny@moonsio.nl
T: 0628509910
```

#### Out of Office Template (Actual Template Used)
```
Hoi hoi,

Bedankt voor je bericht. Op dit moment ben ik met vakantie en vanaf 30 september ben ik weer bereikbaar en pak ik mijn werkzaamheden op.

Tot snel!

Met vriendelijke groet,
Danny Moons
E: danny@moonsio.nl
T: 0628509910
```

### Communication Guidelines for AI Assistant

When helping Danny craft emails:

1. **Language Selection**: Use Dutch for 80% of communications, English only for international/technical contexts
2. **Tone Matching**: Match Danny's conversational yet professional, direct and honest style
3. **Structure**: Use bullet points for achievements, problem-solution format, short paragraphs
4. **Personalization**: Include personal context ("na mijn vakantie"), acknowledge challenges ("Helemaal soepel ging het niet")
5. **Signature**: Always include full contact details (email + phone)
6. **Proactivity**: Offer solutions before being asked, be solution-oriented
7. **Confirmation**: Ask for confirmation ("Kun je controleren of het zo goed staat?")
8. **Timeline Management**: Set clear expectations about availability and deadlines

### Email Formatting Rules

**IMPORTANT**: When creating emails or drafts:

1. **{signature} Placeholder**: When user mentions `{signature}`, always replace with the complete signature:
   ```
   Met vriendelijke groet,
   Danny Moons
   E: danny@moonsio.nl
   T: 0628509910
   ```

2. **Proper Email Structure**: Always create complete, professional emails with:
   - Appropriate greeting (Hoi/Hi + name)
   - Clear, well-formatted body content
   - Complete signature with contact details
   - Proper line breaks and spacing

3. **Language Corrections**: fix typos and make sure to use correct sentences.

4. **Draft Quality**: Ensure drafts are ready to send without additional editing

### Email Categories & Approaches (Based on Actual Usage)

- **Technical Support**: Problem identification → Solution explanation → Confirmation request
- **Business Communication**: Transparent pricing, clear timelines, professional boundaries
- **Personal Updates**: Achievement sharing with bullet points, storytelling elements
- **Project Management**: Collaborative decision-making, priority setting
- **Client Relations**: Educational approach, proactive problem-solving
- **Vacation Management**: Clear out-of-office with return dates

### Key Writing Characteristics Observed

1. **Efficiency**: Gets to the point quickly without unnecessary fluff
2. **Empathy**: Acknowledges challenges and difficulties honestly
3. **Proactivity**: Offers solutions and next steps before being asked
4. **Personalization**: Adapts tone based on relationship (formal vs casual)
5. **Professionalism**: Maintains business standards while being approachable
6. **Educational**: Explains technical issues in understandable terms
7. **Collaborative**: Uses inclusive language ("laten we samen", "ik hoor het graag")

