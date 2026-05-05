# Follow-ups Management System

This folder contains follow-ups that need to be tracked and emailed to owners.

## Structure

- **pending/**: Follow-ups waiting for review and approval
- **approved/**: Follow-ups that have been approved and sent
- **templates/**: Email templates

## File Format

Follow-ups should be JSON files with the following structure:

```json
{
  "follow_ups": [
    {
      "id": "FU-001",
      "action": "Complete project documentation",
      "owner_name": "John Doe",
      "owner_email": "john.doe@example.com",
      "expected_date": "2026-05-15",
      "status": "pending",
      "created_date": "2026-05-05"
    }
  ]
}
```

## Workflow

1. Upload Word files or emails with follow-ups to the `pending/` folder
2. Parse and convert to JSON format
3. Review follow-ups in Power Automate approval flow
4. Approve/Reject each follow-up
5. Approved follow-ups are sent via Outlook and moved to `approved/`
