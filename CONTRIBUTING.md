# Contributing

Thanks for your interest - contributions are welcome.

## Ways to help

- Report a bug or request a feature via [issues](https://github.com/dbhq-uk/outlook-skill/issues)
- Improve the skill instructions or scripts via a pull request

## Local development

```bash
git clone https://github.com/dbhq-uk/outlook-skill.git
cd outlook-skill
./install.sh          # symlinks into ~/.claude/skills (edits are live)
```

Scripts are symlinked, so edits are live immediately. After editing a `SKILL.md`, re-run `./install.sh` to regenerate the installed copy.

## Before opening a PR

- `bash -n skills/*/scripts/*.sh` - scripts parse cleanly
- `claude plugin validate .` - the plugin validates
- Keep credentials out of the repo and out of commits
- British English, plain hyphens, no trailing full stops on headings

## Licence

By contributing you agree your work is licensed under the [MIT licence](LICENSE).
