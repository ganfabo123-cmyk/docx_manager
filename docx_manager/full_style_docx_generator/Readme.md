# Full Style DOCX Generator

A Flask-based API service that processes documents with chaotic styles using LLM to analyze and standardize styles, then generates properly formatted DOCX documents.

## Project Structure

```
full_style_docx_generator/
├── api/
│   └── routes.py         # API endpoints
├── core/
│   ├── document_processor.py  # Core document processing functions
│   └── file_parser.py         # File parsing utilities
├── server.py             # Flask server setup
├── main.py               # Server starter
└── Readme.md             # This documentation
```

## Installation

1. Clone the repository
2. Install dependencies:
   ```bash
   pip install flask python-docx
   ```

## API Endpoints

### POST /parse-txt-file
- **Description**: Downloads and parses a TXT file from a URL
- **Request Body**: `{"url": "http://example.com/file.txt"}`
- **Response**: `{"status": "success", "message": "File parsed and blocks saved"}`
- **Action**: Saves parsed blocks to `data/parsed_blocks.json`

### POST /parse-docx-file
- **Description**: Downloads and parses a DOCX file from a URL
- **Request Body**: `{"url": "http://example.com/file.docx"}`
- **Response**: `{"status": "success", "message": "File parsed and blocks saved"}`
- **Action**: Saves parsed blocks to `data/parsed_blocks.json`

### GET /identify-short-blocks
- **Description**: Identifies short text blocks (< 30 characters) from parsed blocks
- **Response**: `{"short_blocks": [{"id": 1, "content": "Short text"}]}`
- **Action**: Reads blocks from `data/parsed_blocks.json`

### POST /analyze-styles
- **Description**: Accepts style analysis from LLM and saves it
- **Request Body**: `{"styles": [{"id": 1, "type": "h1", "content": "Heading"}]}`
- **Response**: `{"status": "success", "message": "Styles saved successfully"}`
- **Action**: Saves styles to `data/parsed_styles.json`

### POST /generate-document
- **Description**: Generates a DOCX document from parsed styles
- **Request Body**: `{"parsed_styles": [{"id": 1, "type": "h1", "content": "Heading"}]}`
- **Response**: `{"status": "success", "message": "Document generated successfully", "file_path": "path/to/document.docx"}`
- **Action**: Saves generated DOCX to `downloads/generated_document.docx`

## Core Functions

### document_processor.py

1. **parse_document_content(content: str) -> List[Dict[str, Any]]**
   - Parses plain text content into text blocks
   - Returns list of blocks with `id` and `content` fields

2. **identify_short_blocks(blocks: List[Dict[str, Any]]) -> List[Dict[str, Any]]**
   - Identifies short text blocks (content length < 30)
   - Returns filtered list of short blocks

3. **generate_llm_prompt(blocks: List[Dict[str, Any]]) -> str**
   - Generates prompt for LLM to analyze styles
   - Includes instructions for style classification

4. **save_parsed_styles(analysis: List[Dict[str, Any]], output_path: str) -> None**
   - Saves parsed styles to JSON file

5. **generate_styled_content(parsed_styles: List[Dict[str, Any]]) -> str**
   - Generates styled content from parsed styles
   - Returns markdown-formatted content

6. **generate_docx_document(parsed_styles: List[Dict[str, Any]], output_path: str) -> bool**
   - Generates DOCX document from parsed styles
   - Applies appropriate styling for different block types

### file_parser.py

1. **parse_txt_file(file_path: str) -> Optional[List[Dict[str, Any]]]**
   - Parses TXT file into text blocks

2. **parse_docx_file(file_path: str) -> Optional[List[Dict[str, Any]]]**
   - Parses DOCX file into text blocks

3. **parse_file(file_path: str) -> Optional[List[Dict[str, Any]]]**
   - Generic file parser that dispatches based on file extension

## Workflow

1. **File Parsing**: Use `/parse-txt-file` or `/parse-docx-file` to parse documents
2. **Short Block Identification**: Use `/identify-short-blocks` to get short blocks for LLM analysis
3. **Style Analysis**: Send short blocks to LLM for style classification
4. **Style Saving**: Use `/analyze-styles` to save LLM-generated styles
5. **Document Generation**: Use `/generate-document` to create styled DOCX document

## Usage Example

```bash
# Start the server
python main.py

# Parse a TXT file
curl -X POST http://localhost:5000/parse-txt-file -H "Content-Type: application/json" -d '{"url": "http://example.com/sample.txt"}'

# Get short blocks
curl http://localhost:5000/identify-short-blocks

# Analyze styles (simulating LLM response)
curl -X POST http://localhost:5000/analyze-styles -H "Content-Type: application/json" -d '{"styles": [{"id": 1, "type": "h1", "content": "Title"}, {"id": 2, "type": "paragraph", "content": "Content"}]}'

# Generate document
curl -X POST http://localhost:5000/generate-document -H "Content-Type: application/json" -d '{"parsed_styles": [{"id": 1, "type": "h1", "content": "Title"}, {"id": 2, "type": "paragraph", "content": "Content"}]}'
```

## Directory Structure

- **api/**: API routes and endpoints
- **core/**: Core processing functions
- **data/**: Storage for parsed blocks and styles
- **downloads/**: Storage for generated DOCX documents
- **temp/**: Temporary storage for downloaded files

## Error Handling

All endpoints include comprehensive error handling with traceback information for debugging purposes. Error responses include both error message and traceback details.

## Dependencies

- Flask: Web framework
- python-docx: DOCX document generation
- urllib: URL file downloading

## License

This project is intended for internal use within the HIT paper helper system.