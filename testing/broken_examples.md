# Guide for New Contributors

## Handling Known Broken Examples

The project uses a system to track examples that are intentionally broken or contain known issues:

1. Broken examples are listed in `testing/.broken` (one example name per line)
2. The autocheck tool will:
   - Skip broken examples by default
   - Show broken examples with `[BROKEN]` tag
   - Consider broken examples that fail checks as "successful"
   - Alert if a broken example unexpectedly passes all checks

To check all examples including broken ones, use:
```
python3 utils/autocheck.py --all --force
```

To view which examples are listed as broken:
```
python3 utils/autocheck.py --list-broken
```

When an example is fixed, remember to remove it from the `.broken` file. 