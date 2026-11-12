# Contact organisation details

`contact org` and MCP `contact_org` read the Microsoft Graph directory.
They do not change personal contacts or directory records.

## Use the CLI

```bash
outlook-cli contact org me
outlook-cli contact org person@example.com --json
outlook-cli contact org "Alex Example" --chain
outlook-cli contact org person@example.com --chain --max-depth 3 --json
outlook-cli --readonly contact org me
```

The user argument defaults to `me`.
It also accepts an email, UPN or directory user ID.
A name first uses an exact display-name match, then a display-name prefix.
An ambiguous name requires an email or directory user ID.
A cached contact index or personal contact ID does not identify a directory user.

`--chain` adds the recorded management chain, with the immediate manager first.
`--max-depth` requires `--chain` and a positive integer.
A depth limit marks the chain as incomplete because the tool does not check further managers.
Without a depth limit, the tool follows the recorded chain until it ends, becomes unavailable or repeats an ID.
The tool never reads the reports of each direct report recursively.

## Use MCP

```json
{"user":"person@example.com","include_chain":true,"max_depth":3}
```

Call `contact_org` with these arguments.
All arguments are optional.
The user defaults to `me`, and the management chain defaults to off.
The tool remains available under read-only and no-send policies.
Restart the MCP server after an update to expose the new tool.

## Results

| Field | Meaning |
|---|---|
| `source` | Microsoft Graph directory. |
| `person` | The selected person's directory profile. |
| `manager` | The immediate manager, or null when unavailable. |
| `managerStatus` | Available, not set or not visible, or unavailable after an API error. |
| `directReports` | The direct relationships that the directory returns, with one entry for each ID. |
| `directReportsCount` | The number of entries returned, not an assumed company-wide headcount. |
| `directReportsComplete` | False when a page fails or duplicate entries indicate a change during the read. |
| `managementChain` | The requested chain, or null when the caller did not request it. |
| `managementChainComplete` | Whether the tool reached the end of the recorded chain; null when not requested. |
| `warnings` | Explicit information about permission errors, limited profiles, cycles or incomplete results. |

Each profile includes:

- The directory ID and display name.
- The email and UPN.
- The job title and department.
- The company name.
- The object type and a flag for limited information.

Missing profile fields remain null.
The CLI labels them as not recorded or visible.
An entry with only an ID remains visible and has `limitedInformation=true`.
The tool does not infer a manager from email conversations, job titles or company names.

Direct reports use every response page.
The page size of 100 is not a total result limit.
Name lookup reads at most two candidates to detect ambiguity, not to list all name matches.
A continuation link also prevents the tool from treating one returned name as a unique match.

A partial report list retains the entries that the tool already read.
For example, zero returned entries with `directReportsComplete=false` does not mean that the person has no reports.
An expired token stops the read rather than producing an empty result.

The chain ends where the available directory records end.
This does not prove that its last entry is the company's highest manager.
An organisational contact can appear as a manager, but the tool does not follow that contact through the user endpoint.

## Access and limitations

Organisation details require a work or school directory account.
External people may have no accessible directory profile.
A guest record may contain less information than an internal user record.

Microsoft documents `User.Read.All` as the least delegated permission for the manager endpoint.
The direct-reports endpoint documents `User.Read` with `User.ReadBasic.All` as its least delegated permissions.
Other directory permissions can also grant access.
The account's existing token may already provide it.
The tool does not request new permissions or change authentication settings.

The directory can return restricted entries with only an ID and object type.
The tool preserves that distinction instead of inventing names or organisation details.
A missing manager can mean that no relationship is recorded or visible.
HTTP 403 produces an explicit warning rather than a claim that no manager or reports exist.

The organisation can maintain incomplete or outdated records.
This feature reports those records, not an independently verified company structure.

## Implementation and verification

`outlook_draft/directory.py` contains the shared Graph methods and result format.
The CLI and MCP use the same result helper.
All directory requests use GET.
Pagination rejects a different host, a different relationship or a repeated page URL.

Synthetic tests cover resolution, pagination, permissions, partial results and the optional management chain.
Live read-only checks verified a profile, its manager and another user's direct reports with the current account.
No directory records or contacts changed.

Microsoft references:

- [Get a user's manager](https://learn.microsoft.com/en-us/graph/api/user-list-manager?view=graph-rest-1.0).
- [List direct reports](https://learn.microsoft.com/en-us/graph/api/user-list-directreports?view=graph-rest-1.0).
