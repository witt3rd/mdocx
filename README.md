# mdocx

[![Rye](https://img.shields.io/endpoint?url=https://raw.githubusercontent.com/astral-sh/rye/main/artwork/badge.json)](https://rye.astral.sh)
[![GitHub license](https://img.shields.io/github/license/witt3rd/mdocx.svg)](https://github.com/witt3rd/mdocxe/blob/main/LICENSE)
[![GitHub issues](https://img.shields.io/github/issues/witt3rd/mdocx.svg)](https://github.com/witt3rd/mdocx/issues)
[![GitHub stars](https://img.shields.io/github/stars/witt3rd/mdocx.svg)](https://github.com/witt3rd/mdocx/stargazers)
[![Twitter](https://img.shields.io/twitter/url/https/twitter.com/dt_public.svg?style=social&label=Follow%20%40dt_public)](https://twitter.com/dt_public)

mdocx is a Python tool that converts Markdown files to DOCX format with custom styling.

## Installation

This project uses [rye](https://rye.astral.sh/) for dependency management. To set up the project:

1. Ensure you have rye installed. If not, follow the [rye installation guide](https://rye.astral.sh/guide/installation/).

2. Clone the repository:

   ```sh
   git clone https://github.com/yourusername/mdocx.git
   cd mdocx
   ```

3. Install dependencies:

   ```sh
   rye sync
   ```

## Usage

To use mdocx, run the following command:

```sh
rye run python src/mdocx/main.py <input_markdown> <template_docx> <output_docx>
```

Where:

- `<input_markdown>` is the path to your input Markdown file
- `<template_docx>` is the path to your template DOCX file (containing the desired styles)
- `<output_docx>` is the path where you want to save the output DOCX file

For example:

```
rye run python src/mdocx/main.py input.md template.docx output.docx
```

## Features

- Converts Markdown to DOCX format
- Supports custom styling based on a template DOCX file
- Handles various Markdown elements including headings, paragraphs, lists, code blocks, and inline styles

## Dependencies

- python-docx
- mistletoe

## Development

To set up the development environment:

1. Clone the repository
2. Run `rye sync` to install dependencies
3. Make your changes
4. Run tests (if available)

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

[Add your chosen license here]

## Contact

Donald Thompson - <witt3rd@witt3rd.com>

Project Link: [https://github.com/witt3rd/mdocx](https://github.com/witt3rd/mdocx)
