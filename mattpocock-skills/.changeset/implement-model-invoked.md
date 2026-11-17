---
"mattpocock-skills": minor
---

`implement` is model-invoked, so the agent reaches for it when asked to implement, build or fix a named tracker issue or scope, and other skills can hand it an issue to build: `disable-model-invocation` and the Codex `policy` block are removed, its description is model-facing, and both READMEs list it under Model-invoked. It takes a ticket reference from the skill that called it as well as from the user. When the session already holds a rewind point, set by the user or by the calling skill, it keeps that point if it clearly still holds and builds in one part, and otherwise asks the user whether to keep it or set a new one.
