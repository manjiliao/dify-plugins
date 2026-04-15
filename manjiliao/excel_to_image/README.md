# Dify FreeSpire Excel To Image Plugin

This is a local Dify tool plugin that converts Excel worksheets into image files by using `FreeSpire.XLS for Python`.

## Features

- Accepts uploaded Excel files directly from Dify
- Supports `xls`, `xlsx`, `xlsm`, `xlsb`, and `ods`
- Exports all sheets or selected sheets
- Supports `png` and `jpg`
- Optionally removes worksheet page margins before rendering

## Contact & Repository

- **Author**: manjiliao
- **Repository**: https://github.com/manjiliao/dify-plugins
- **Contact**: [Your contact email or method]

## Privacy Policy

This plugin follows Dify's privacy protection guidelines. For detailed privacy policy, please refer to [PRIVACY.md](PRIVACY.md).

## Project Structure

```text
.
|-- manifest.yaml
|-- main.py
|-- provider/
|   |-- excel_to_image.py
|   `-- excel_to_image.yaml
|-- tools/
|   |-- excel_to_image.py
|   `-- excel_to_image.yaml
`-- requirements.txt
```

## Install Dependencies

Use Python `3.11+` first, then install:

```powershell
py -3.11 -m venv .venv
.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

## Local Debug

1. In Dify, open `Plugins` and start remote debugging.
2. Copy `.env.example` to `.env`.
3. Fill in the debug address and key.
4. Install dependencies.
5. Run:

```powershell
python -m main
```

If you install dependencies into the project-local `.packages` directory like I did in this environment, use:

```powershell
powershell -ExecutionPolicy Bypass -File .\run-local.ps1
```

## Tool Parameters

- `excel_content`: uploaded Excel file
- `sheets`: `all` or a comma-separated list such as `Sheet1,2,4-6`
- `image_format`: `png` or `jpg`
- `remove_margins`: whether to clear page margins before rendering

## Packaging

After the plugin works in remote debug mode:

```powershell
dify plugin package .
```

The generated artifact is a `.difypkg` file that you can upload into Dify.

## Notes

- `requirements.txt` is pinned to `Spire.Xls.Free==14.12.4`, which is the free-edition package line.
- The free edition may add evaluation watermarks or impose export limitations.
- In this workspace I verified the rendering library by generating a sample file at `_outputs/spire_test.png`.
- Some complex Excel layouts may render differently from Microsoft Excel.

## Publishing to Dify Marketplace

To publish this plugin to the Dify Marketplace:

### Development Checklist
- [x] Plugin developed and tested according to [Plugin Developer Guidelines](https://docs.dify.ai/en/develop-plugin/publishing/standards/contributor-covenant-code-of-conduct)
- [x] Privacy Policy created in `PRIVACY.md`
- [x] Privacy policy path included in `manifest.yaml`
- [x] Contact information and repository URL added to README.md

### Publishing Steps

1. **Package the plugin**:
   ```powershell
   dify plugin package .
   ```
   This creates a `.difypkg` file for distribution.

2. **Fork the repository**:
   Fork [dify-plugins repository](https://github.com/langgenius/dify-plugins/fork)

3. **Create directory structure**:
   - Create an organization directory (e.g., `manjiliao`)
   - Create a plugin subdirectory (e.g., `manjiliao/excel_to_image`)
   - Place source code and `.difypkg` file in that subdirectory
   - Example: `manjiliao/excel_to_image/excel_to_image-0.1.0.difypkg`

4. **Submit Pull Request**:
   - Follow the required PR template format
   - Wait for review

5. **After approval**:
   - Plugin code merges into main branch
   - Plugin automatically listed on [Dify Marketplace](https://marketplace.dify.ai/)

### Updating the Plugin

When releasing updates:

1. Increment version in `manifest.yaml`
2. Each PR should contain only one file change - the new `.difypkg` file
3. Verify the version hasn't been published before
4. Document breaking changes clearly in README.md
5. Consider using [GitHub Actions workflow template](https://docs.dify.ai/en/develop-plugin/publishing/marketplace-listing/plugin-auto-publish-pr) for automated PR creation

## Security

For security issues, please contact [security@dify.ai](mailto:security@dify.ai) instead of posting on GitHub.
