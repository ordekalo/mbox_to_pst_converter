# MBOX to PST Converter

This project is a simple Python script to convert email messages from an MBOX file format to a PST file format using `pypff` (Python bindings for the `libpff` library) and Python's `mailbox` module.

## Prerequisites

- Python 3.x installed on your machine.
- Install dependencies by running:
  
  ```bash
  pip install -r requirements.txt
  ```

## How to Use

1. Place your MBOX file in a known directory.
2. Modify the paths in the `mbox_to_pst_converter.py` file for your MBOX and desired PST file locations.

   ```python
   mbox_file = "path/to/your/file.mbox"
   pst_file = "path/to/your/file.pst"
   ```

3. Run the script:

   ```bash
   python mbox_to_pst_converter.py
   ```

4. After completion, you will see the `Conversion completed!` message and your PST file will be generated.

## Dependencies

- `pypff`: A library for reading and writing Personal Folder File (PST) files.
- Python `mailbox`: A module from the Python standard library used to read MBOX files.

To install dependencies, run:

```bash
pip install -r requirements.txt
```

## Notes

- **Compatibility**: This script assumes basic MBOX messages. Complex MBOX structures (such as HTML messages or attachments) may require further processing.
- **Error Handling**: The script does not handle all edge cases like corrupt MBOX files or malformed messages. It can be extended to do so.
- **PST Support**: The `pypff` library has limited support and is not actively maintained, so it may not work for all versions of PST files. For more robust solutions, consider commercial libraries like `Aspose.Email`.

## License

This project is licensed under the MIT License.
