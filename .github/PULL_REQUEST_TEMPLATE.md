## What changed
<!-- One or two sentences describing the fix or feature. -->

## Automated checks
- [ ] CI tests pass (`node scripts/run-tests.js`)

## Manual QA — only required if this PR touches CSS or rendering patches

If your change affects **signatures, display names, reading pane layout, or compose view**, load the extension in Safari and verify:

**Reading pane**
- [ ] Sender display name shows correctly (First Last, no middle name bleed)
- [ ] Signature block renders without extra whitespace or broken layout
- [ ] Long department lines wrap cleanly and don't overflow

**Compose view**
- [ ] Signature appears in the correct position
- [ ] Existing email body is not displaced or duplicated

**General**
- [ ] No console errors in Safari Web Inspector
- [ ] Patch does not fire on unrelated pages (check `urlPattern`)

<!-- If none of the above apply, delete this section. -->
