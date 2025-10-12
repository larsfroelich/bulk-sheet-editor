# Bulk sheet editor

## Description
Simple GUI application for bulk-creating sheets (tables/documents) by modifying individual elements of a template
using input from a csv-file.

## Features
> WARNING - This is a work in progress. Use at your own risk.
- import dats using CSV
- import Open File Format (ODF) files (sheets or docs)
- map csv-columns to template-elements (cells or placeholders)
- bulk create pages/sheets/documents

## Installation
For now, Bulk sheet editor is still in development and there is no initial release yet.
To build and install directly from source, you can use

`cargo install --git https://github.com/larsfroelich/file-kraken.git`

this requires installing [Rust/cargo](https://rust-lang.org/tools/install/):

`curl --proto '=https' --tlsv1.2 -sSf https://sh.rustup.rs | sh`
