# Release and Build Instructions

## Automated GitHub Actions Release

This project is configured to automatically build and release executables via GitHub Actions.

### How it Works

1. **Automatic Builds on Tags**: When you push a version tag (e.g., `v1.0.0`), GitHub Actions automatically:
   - Builds the Windows executable
   - Creates a GitHub Release
   - Attaches the `.exe` file to the release

2. **Manual Trigger**: You can also manually trigger a build from the GitHub Actions tab

### Creating a New Release

Follow these steps to create a new release:

```bash
# 1. Commit all your changes
git add .
git commit -m "Release version 1.0.0"

# 2. Create and push a version tag
git tag v1.0.0
git push origin main
git push origin v1.0.0

# 3. GitHub Actions will automatically:
#    - Build the executable
#    - Create a release on GitHub
#    - Upload the .exe file
```

### Version Tag Format

- Use semantic versioning: `vMAJOR.MINOR.PATCH`
- Examples: `v1.0.0`, `v1.2.3`, `v2.0.0-beta`

### Viewing Releases

1. Go to your GitHub repository
2. Click on "Releases" in the right sidebar
3. Download the `.exe` file from the latest release

---

## Local Build (For Testing)

If you want to build the executable locally before creating a release:

### Quick Build (Windows)

Simply run the batch file:
```bash
build_exe.bat
```

The executable will be created at: `dist\FileToDBConverter.exe`

### Manual Build

```bash
# Install PyInstaller (if not already installed)
uv pip install --system pyinstaller

# Clean previous builds
rmdir /s /q build dist

# Build the executable
pyinstaller FileToDBConverter.spec

# Your executable is now at: dist\FileToDBConverter.exe
```

---

## Build Configuration

The build is configured via `FileToDBConverter.spec`:

- **Single File**: Creates a standalone `.exe` (no folder required)
- **No Console**: Runs as a GUI application (no terminal window)
- **All Dependencies**: Includes Python runtime and all libraries
- **Hidden Imports**: Ensures all required modules are included

### Customization

To modify the build:

1. Edit `FileToDBConverter.spec`
2. Common changes:
   - **Add an icon**: Set `icon='path/to/icon.ico'`
   - **Show console**: Change `console=False` to `console=True`
   - **Include data files**: Add to `datas` list

---

## Testing Your Release

Before creating a GitHub release:

1. **Build locally**: Run `build_exe.bat`
2. **Test the exe**: Run `dist\FileToDBConverter.exe`
3. **Verify functionality**:
   - Can add files
   - Preview dialog works
   - Database connection works
   - Data import succeeds
4. **Test on clean machine**: Copy the `.exe` to a machine without Python installed

---

## Troubleshooting

### Build Fails with Missing Module

Add the module to `hiddenimports` in `FileToDBConverter.spec`:
```python
hiddenimports=[
    'tkinter',
    'your_missing_module',
    # ... other imports
],
```

### Antivirus Flags the Executable

This is common with PyInstaller builds. To resolve:
1. Submit to antivirus vendors as false positive
2. Sign the executable with a code signing certificate (advanced)

### Large File Size

The `.exe` includes Python runtime and all dependencies. Expected size: 40-80 MB

To reduce size:
- Remove unused dependencies from `requirements.txt`
- Use UPX compression (already enabled in spec file)

---

## GitHub Actions Workflow

The workflow file is located at: `.github/workflows/build-release.yml`

### Workflow Triggers

- **Tag Push**: Automatically builds when you push a tag starting with `v`
- **Manual**: Can be triggered from GitHub Actions tab

### Workflow Steps

1. Checkout code
2. Set up Python 3.11
3. Install uv and dependencies
4. Build executable with PyInstaller
5. Create release archive (ZIP)
6. Upload as artifact (available for 30 days)
7. Create GitHub Release (if triggered by tag)

### Viewing Build Logs

1. Go to "Actions" tab on GitHub
2. Click on the workflow run
3. View logs for each step

---

## Release Checklist

Before creating a release:

- [ ] All features tested and working
- [ ] Version number updated in code (if applicable)
- [ ] CHANGELOG updated with new features
- [ ] Local build succeeds
- [ ] Executable tested on clean machine
- [ ] Commit all changes to main branch
- [ ] Create and push version tag
- [ ] Verify GitHub Actions build succeeds
- [ ] Test downloaded release from GitHub

---

## Future Improvements

Consider adding:
- [ ] macOS build support (requires macOS runner)
- [ ] Auto-update functionality
- [ ] Code signing certificate
- [ ] Installer (e.g., InnoSetup)
- [ ] Portable version (ZIP with configs)
