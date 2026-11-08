# Security: Tableau user logic in Power BI

Tableau `USERNAME()`, `FULLNAME()`, `ISMEMBEROF()` and user filters cannot be translated one to one. The tools therefore stop the workflow (`BLOCKED`) until an architect approves the design.

- **Identity.** `USERNAME()` is a Tableau login; Power BI's `USERPRINCIPALNAME()` is the Entra sign-in (usually an email). They match only through a user table that maps one to the other.
- **Per-viewer rows.** A calculation that compares `USERNAME()` with a column changes per viewer. In Power BI this is a row-level security (RLS) role on the semantic model, not a calculated column (Desktop refuses `USERPRINCIPALNAME()` in one). The build uses a single `preview_login` (decisions.json) so the report opens; that is a preview, not security.
- **Hard-coded logins** (admin lists) move into the user/security table, never into measures.
- **Define RLS once**, in the shared semantic model, and test it with *View as* for at least one user per role and one user with no rows.
- **Approval.** Record it in `input\decisions.json`: `"workflow": {"rls_approved_by": "<architect name>"}`. The tools never write this for you.
