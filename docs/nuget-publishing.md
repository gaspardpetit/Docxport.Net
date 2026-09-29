# NuGet Trusted Publishing

The NuGet release workflow uses GitHub OIDC to obtain a short-lived publishing
key. Configure the matching policy as described in [NuGet.org Trusted Publishing](https://learn.microsoft.com/en-us/nuget/nuget-org/trusted-publishing)
before running the workflow:

| Policy field | Value |
| --- | --- |
| Package owner | `gaspardpetit` |
| GitHub repository owner | `gaspardpetit` |
| GitHub repository | `Docxport.Net` |
| Workflow file | `publish-packages.yml` |
| Environment | Leave empty; this job does not use a GitHub environment. |
| Package scope | Allow new versions of both `DocxportNet` and `DocxportNet.Cli`. |

Enter only the workflow filename, without `.github/workflows/`. The workflow
requests an OIDC token immediately before the push, then uses the short-lived
key returned by `NuGet/login` for both packages. No GitHub API-key secret is
needed.

To publish the missed `1.7.0` packages after this workflow reaches `main`, run
**Publish NuGet Packages** with **Run workflow** on `main` and `version=1.7.0`.
Rerunning the failed release run would use its original workflow revision.
The push step skips packages whose version is already present on NuGet.org.
After a successful publish, remove the obsolete `NUGET_API_KEY` repository
secret.
