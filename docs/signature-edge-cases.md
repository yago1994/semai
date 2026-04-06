# Signature Edge Cases

Track signature-detection failures and the heuristic added for each case.

## Template

- Date:
- Surface:
  - Regular view, chat view, or both
- Sender / context:
- Sample structure:
- Failure mode:
- Heuristic added or proposed:
- Result:

## Cases

### Daniel Drane / Emory contact card

- Surface:
  - Chat view preserved the full signature block while regular view hid it
- Sample structure:
  - Message text
  - Short signoff line with the sender's name (`daniel`)
  - Nearby repeated sender name (`Daniel L. Drane, Ph.D., ABPP(CN)`)
  - Followed by dense title / institution / address / phone / email lines
- Failure mode:
  - No stable Outlook signature class or id
  - Original clone-cleaning path kept the contact card because the live-view-only heuristic caught it first
- Heuristic added:
  - Clone-side rule that strips trailing blocks when a name repeats in close proximity and the following block contains 2+ contact signals
- Result:
  - Intended to keep the message body while removing the trailing contact-card signature in chat view

### Mandi Schmitt / UserTesting outreach signature

- Surface:
  - Chat view preserved a nested table-based signature and opt-out footer
- Sample structure:
  - Outreach copy ending with `Best,` and `Mandi`
  - Followed by a table-based signature card with headshot/logo, role, email, LinkedIn, and Calendar links
  - Followed by a trailing opt-out link / tracking footer
- Failure mode:
  - Signature content was nested several levels deep inside tables and wrapper divs, so shallow sibling heuristics missed it
- Heuristic added:
  - Descendant-based repeated-name detection in the trailing portion of the clone
  - Requires nearby repeated name tokens plus 2+ contact signals and at least one compact contact-card-like block
- Result:
  - Intended to remove the signature card and trailing outreach footer while preserving the actual message body

### Leah Ekube / Pendo compact branded signature

- Surface:
  - Signature still visible in chat view
- Sample structure:
  - Compact nested `div[dir="ltr"]` signature-only block
  - Name with pronouns
  - Title plus branded domain link
  - Separator line and branded image banner
- Failure mode:
  - No obvious Outlook signature class or id
  - Signature is compact and image-heavy rather than a long contact card
- Heuristic added or proposed:
  - Track as a dedicated compact branded-signature case
  - Likely needs a clone-side rule for short signature-only blocks with name/title/domain plus separator/image
- Result:
  - Pending implementation
