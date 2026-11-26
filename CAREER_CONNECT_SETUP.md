# Career Connect setup

Career Connect is a static Astro page backed by a separate Supabase project. Without the configuration below, it displays an unavailable message and does not accept submissions. Profiles and recommendations are public only after an administrator approves them; email addresses remain private. The site does not collect CVs.

## Before accepting real submissions

- The example policies allow anonymous submissions. They do not provide CAPTCHA, rate limiting, identity verification, or abuse prevention. Add these protections using a trusted server or edge function before sharing the live submission form publicly. Browser checks are not a security boundary.
- Decide and publish a privacy notice and retention period suitable for your jurisdiction, including the storage provider, who can access information, handling of recommendations, and how people can request access, correction, and deletion. The short notice on the form is not a complete privacy policy.
- Recommendations are unverified claims. Review them, including whether the author has permission to name the candidate, before approving them. Do not rely on them as verified references.
- Set up a process to review pending submissions and remove profiles and recommendations on request.
- Restrict access to the Supabase project and use appropriate account security. Never put a service-role key in the site or in a public environment variable.

## Configure

1. Create a Supabase project and run `supabase/career-connect.sql` once in its SQL editor. The SQL creates two tables and public read-only views. It expects a fresh setup; it is not a repeatable migration. Do not run it during deployments or on an existing database.
2. Set `PUBLIC_CAREER_SUPABASE_URL` to the project's HTTPS API URL and `PUBLIC_CAREER_SUPABASE_ANON_KEY` to its publishable/anon key in your build environment. These values are public in the generated site. Never use a secret or service-role key here.
3. Rebuild and deploy the site. Public environment values are embedded at build time.
4. Verify with a test submission: the profile and recommendation should appear only after approval. Confirm that anonymous users cannot read the underlying tables or private email.

## Website updates and database safety

The website's build generates static files and does not run the SQL setup or alter the database. Profiles and recommendations live in the separate Supabase project, not in the repository or the build output. Copying the repository and rebuilding or replacing the website files does not erase that data.

For every deployment, keep using the same Supabase project URL and public key in the build environment. A different project will show a different directory; missing settings will make the directory unavailable. Do not delete or recreate the Supabase project, run the initial setup again, or add database reset commands to the deployment process. Normal submissions and administrator actions can still change data independently of a website build.

Maintain backups and test restoration before making database changes. Check the backup and retention options available on your Supabase plan, and make an additional backup before any manual schema migration. Schema changes must be reviewed and applied separately; copying a new SQL file into the repository does not apply it to Supabase.

## Existing CV-enabled installations

If the earlier version of this setup was already applied, removing the form does not remove its existing storage policy. To stop further anonymous CV uploads, run the following once in the Supabase SQL editor for that existing project:

```sql
drop policy if exists "Upload CVs for review" on storage.objects;
revoke insert (cv_path) on public.career_profiles from anon, authenticated;
```

This does not delete the old bucket, files, `cv_path` column, profiles, or recommendations. Review any existing files and your privacy obligations before deleting them through Supabase; do not assume rebuilding the website removes them. New installations do not create the bucket or column.

## Moderation and deletion

In the Supabase dashboard, use the table editor to inspect `career_profiles` and `career_endorsements`. Set `approved` to `true` only after review. Recommendations appear only when both the recommendation and its candidate are approved. Only authorized administrators should have dashboard access.

To remove a candidate, delete their profile row. Associated recommendations are deleted automatically. Delete recommendations separately when requested. If an earlier CV-enabled installation has stored files, a database row deletion does not delete those files; handle them separately according to your privacy obligations.