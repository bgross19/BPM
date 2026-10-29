# BPM Project Rules & GitHub Workflow

## GitHub Delivery Rule: Always Use Branches & Pull Requests
- **Never push commits directly to `main`** unless explicitly instructed by the user.
- Whenever pushing code changes to `https://github.com/bgross19/BPM`:
  1. Validate JavaScript syntax in `Code.gs` and `JavaScript.html` using WebKit's built-in JavaScriptCore (`/System/Library/Frameworks/JavaScriptCore.framework/Versions/Current/Helpers/jsc`).
  2. Create a descriptive feature branch (e.g., `ux/mobile-and-admin-improvements` or `fix/...`) from the base commit SHA via the GitHub Git Data REST API (`POST /repos/bgross19/BPM/git/refs` with `ref: "refs/heads/<branch>"`).
  3. Commit the updated tree to that feature branch (`POST /repos/bgross19/BPM/git/trees` -> `POST /repos/bgross19/BPM/git/commits` -> `PATCH /repos/bgross19/BPM/git/refs/heads/<branch>`).
  4. Open a Pull Request on GitHub (`POST /repos/bgross19/BPM/pulls` with `head: "<branch>"` and `base: "<base-branch>"`) with a clear title and structured summary of the changes, and share the PR link with the user.
