// ================================================================
// job/directoryService.js
// ----------------------------------------------------------------
// PURPOSE : Read the company user list from Microsoft 365 so tracker names
//           can be matched to email addresses. Only reads (SRP); filtering
//           is recipientResolver.directoryFromGraphUsers, matching is
//           RecipientResolver.
//           Equivalent of Excel_Notify\Services\GraphDirectoryService.cs,
//           but app-only (User.Read.All, granted by IT 2026-10-13).
//
// CHANGE LOG
// - 2026-10-14 : Initial version.
//     * Whole list fetched ONCE per run (~135 users) and matched locally
//       (same reason as C#: one call, no throttling).
//     * Follows @odata.nextLink pages (Graph returns at most 999 per page).
//     * accountEnabled is also read, so staff who left (disabled account)
//       are never emailed - promised in the IT checklist, Task 7.
// ================================================================

import { directoryFromGraphUsers } from '../core/recipientResolver.js';

const FIRST_PAGE = 'users?$select=displayName,mail,userPrincipalName,accountEnabled&$top=999';

/** @param {import('./graphClient.js').GraphClient} graph */
export async function getCompanyDirectory(graph) {
  const users = [];
  let next = FIRST_PAGE;
  while (next) {
    const page = await graph.getJson(next);
    users.push(...(page.value || []));
    next = page['@odata.nextLink'] || null; // absolute URL - graphClient accepts it
  }
  return directoryFromGraphUsers(users);
}
