# Project Structure Reorganization

## Current Structure (Flat)
```
file_to_database_table/
├── gui.py
├── main.py
├── requirements.txt
├── FileToDBConverter.spec
├── build_exe.bat
├── README.md
├── RELEASE.md
├── config.json (generated)
├── .encryption_key (generated)
├── logs/ (generated)
├── build/ (generated)
├── dist/ (generated)
└── .github/
    └── workflows/
        └── build-release.yml
```

## Proposed Structure (Organized)
```
file_to_database_table/
├── src/                          # Source code
│   ├── __init__.py
│   ├── gui.py                    # GUI application
│   ├── database.py               # Database operations
│   ├── file_processor.py         # File reading and processing
│   ├── utils.py                  # Utility functions (sanitize, encryption)
│   └── dialogs/                  # GUI dialog components
│       ├── __init__.py
│       ├── preview_dialog.py     # Data preview dialog
│       └── connection_dialog.py  # Connection manager dialog
│
├── scripts/                      # Build and utility scripts
│   ├── build_exe.bat            # Windows build script
│   └── build_exe.sh             # Linux/Mac build script (future)
│
├── docs/                         # Documentation
│   ├── README.md                # Main documentation
│   ├── RELEASE.md               # Release guide
│   ├── CONTRIBUTING.md          # Contribution guidelines
│   └── CHANGELOG.md             # Version history
│
├── tests/                        # Unit tests (future)
│   ├── __init__.py
│   ├── test_database.py
│   └── test_file_processor.py
│
├── resources/                    # Resources (future)
│   ├── icon.ico                 # Application icon
│   └── sample_files/            # Sample CSV/Excel files
│
├── .github/                      # GitHub configuration
│   └── workflows/
│       └── build-release.yml    # CI/CD workflow
│
├── requirements.txt              # Python dependencies
├── FileToDBConverter.spec        # PyInstaller config
├── .gitignore
├── LICENSE
│
├── config.json                   # Generated - Database configs
├── .encryption_key               # Generated - Encryption key
├── logs/                         # Generated - Application logs
├── build/                        # Generated - Build artifacts
└── dist/                         # Generated - Executables
```

## Benefits

### 1. **Separation of Concerns**
- Source code in `src/`
- Documentation in `docs/`
- Build scripts in `scripts/`
- Tests in `tests/`

### 2. **Modular Code**
- Split monolithic `main.py` into:
  - `database.py` - Database operations
  - `file_processor.py` - File reading
  - `utils.py` - Helper functions
- Split `gui.py` dialogs into separate modules

### 3. **Easier Navigation**
- Clear hierarchy
- Related files grouped together
- Less clutter in root directory

### 4. **Better Testing**
- Dedicated `tests/` directory
- Easier to add unit tests
- CI/CD integration ready

### 5. **Professional Appearance**
- Industry-standard structure
- Easier for contributors
- Better for open-source projects

## Migration Steps

1. Create directory structure
2. Move files to new locations
3. Update imports in Python files
4. Update PyInstaller spec
5. Update GitHub Actions paths
6. Test build process
7. Update documentation

## Breaking Changes

### For Developers
- Import paths change from `from main import X` to `from src.database import X`
- Build command stays the same (handled by spec file)

### For End Users
- **No changes** - Executable works exactly the same
- Config file location unchanged
- No migration needed

## Rollback Plan

If issues arise:
1. Git revert to previous commit
2. Old flat structure remains in git history
3. Can revert instantly with: `git reset --hard HEAD~1`
