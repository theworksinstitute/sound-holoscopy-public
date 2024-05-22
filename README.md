# Sound Holoscopy

### Contents

- `desktop_ui` for selecting files and parameters for analysis
- `singnal_processing` audio file and signal functions
- `sheet_export` module that uses signal processing functions to create spreadsheet from a template

- `dash_interface` - discontinued

### Project Resources
- [Notes from Paul](https://docs.google.com/document/d/1euY72fLkAkhHFPLVxo7MC_5qruidC_FB0oS_fwit7O4/edit)
- [Sample output](https://docs.google.com/spreadsheets/d/1Ql4yyEOivMTl39vfTZp83TRvd38VV8aduF7trQIAHkk/edit#gid=835315956)
- [sample files](https://drive.google.com/drive/u/3/folders/1cYAhQJw4wVG8jYjvwULwOj-gEGJND2uG)


## Development
### Environment
Standard python evironment. See [wiki](https://github.com/theworksinstitute/wiki) for instructions for setup

### Notes
#### Getting small slice of audio for tests
The script below cuts the audio file in provided path, so that only small beginning portion can be used in tests.

```sh
python test/create_sample.py $FILE_NAME
```
