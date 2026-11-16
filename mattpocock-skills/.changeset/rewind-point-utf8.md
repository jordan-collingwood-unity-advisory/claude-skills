---
"mattpocock-skills": patch
---

The rewind point script prints in UTF-8 whatever the console's default, so a user message holding a character outside a Windows console's legacy code page, such as `❯`, no longer stops it with a `UnicodeEncodeError`.
