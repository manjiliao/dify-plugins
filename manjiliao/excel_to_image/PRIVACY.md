# Privacy Policy

This plugin converts uploaded Excel files into image files by using `FreeSpire.XLS for Python`. This document explains what data the plugin processes and how that data is handled.

## Data Processing

- **Uploaded Excel Files**: The plugin receives Excel files provided by the user at runtime, including formats such as `xls`, `xlsx`, `xlsm`, `xlsb`, and `ods`.
- **Worksheet Rendering**: The plugin reads workbook content, selects one or more worksheets, and renders the selected worksheet content into `png` or `jpg` image files.
- **Export Parameters**: The plugin processes user-supplied export options such as selected sheets, image format, margin removal, scale percentage, and output DPI.
- **Generated Images**: The generated worksheet images are returned to the Dify workflow or tool execution result.

## Data Storage

- **No Persistent File Storage**: The plugin does not intentionally store uploaded Excel files or generated images as permanent application data.
- **Temporary Processing Only**: Uploaded files and output images are written to temporary working directories only for the duration of a tool invocation and are removed after processing completes.
- **No Built-in Database Storage**: The plugin does not write user files, workbook contents, or generated images to a database.

## Logging and Metadata

- **Operational Logs**: The plugin may write runtime logs for troubleshooting and execution tracing.
- **Logged Metadata**: These logs may include limited metadata such as file names, selected sheet options, image format, scale percentage, DPI settings, and Dify runtime identifiers such as `user_id`, `conversation_id`, `app_id`, or `message_id`.
- **No Intentional Content Logging**: The plugin is not designed to log full workbook contents or generated image binary data.

## Third-Party Services

- **No External Content Processing Service**: The plugin does not send Excel file contents or generated images to OpenRouter, OpenAI, or any other third-party AI or image-generation service.
- **Local Library Dependency**: The plugin uses the `FreeSpire.XLS` library locally to load workbook files and render worksheet images.
- **Dify Platform Handling**: When the plugin is used within Dify, uploaded files, invocation metadata, and returned results are handled through the Dify platform and its configured infrastructure.

## Data Retention

- Uploaded Excel files are processed only during the active invocation.
- Generated image files are created temporarily and returned as tool output.
- The plugin itself does not intentionally retain workbook content or generated image content after processing completes.

## Data Transmission

- The plugin does not require any external API call to convert Excel files into images.
- Data transmission related to plugin execution occurs only within the Dify plugin runtime environment and the Dify platform used to invoke the plugin.
- The plugin does not intentionally share user content with unrelated third parties.

## Security Notes

- If remote debugging is enabled, plugin execution metadata and logs may be visible within the Dify debugging environment.
- The free edition of `FreeSpire.XLS` may apply evaluation limitations or watermarks depending on the input file and library behavior.
