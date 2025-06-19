# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview
This is a Google Apps Script (GAS) web application for a library management system (図書館管理システム). The system is written in Japanese and handles book lending, returning, and registration operations.

## Development Commands

### CLASP Commands
- Deploy to Google Apps Script: `clasp push` or `npx clasp push`
- Download from Google Apps Script: `clasp pull` or `npx clasp pull`
- Open in Apps Script Editor: `clasp open` or `npx clasp open`
- View logs: `clasp logs` or `npx clasp logs`
- Run a function: `clasp run <functionName>` or `npx clasp run <functionName>`

Note: If `clasp` command doesn't work directly, use `npx clasp` instead, or set up the npm global path:
```bash
export PATH=~/.npm-global/bin:$PATH
```

## Architecture

### Data Storage
The application uses Google Sheets as the database with the following sheets:
- **書籍DB** (Book Database): Stores book information with columns A:書籍ID (Book ID), B:書籍名 (Book Title)
- **利用者DB** (User Database): Stores user information with columns A:利用者ID (User ID), B:利用者名 (User Name), C:メールアドレス (Email)
- **貸出記録** (Lending Records): Tracks lending/return status with columns:
  - A:書籍ID (Book ID)
  - B:書籍名 (Book Title)
  - C:利用者ID (User ID)
  - D:利用者名 (User Name)
  - E:貸出日時 (Lending Date)
  - F:返却予定日 (Due Date)
  - G:返却状況 (Return Status) - Values: "未返却" (Not Returned) or "返却済み" (Returned)
  - H:返却日時 (Return Date)

### Routing System
The web app uses URL parameters for page routing:
- Default: Menu page (メニュー画面) - Shows navigation to all features
- `?page=checkout`: Book lending system (図書貸出システム)
- `?page=return`: Book return system (図書返却システム)
- `?page=finder`: Rental books search (貸出書籍検索システム)
- `?page=user_returns`: User-based returns (利用者別返却システム)
- `?page=register`: Book registration (書籍登録システム)

### Key Functions in コード.js
- `doGet(e)`: Main entry point that routes to appropriate HTML pages
- `getBookDetails(bookId)`: Retrieves book information from 書籍DB
- `getUserInfo(userId)`: Retrieves user information from 利用者DB
- `getLendingInfo(bookId)`: Finds active lending records for a book
- `processReturnForm(bookId)`: Processes book returns
- `lendBook(bookId, userId)`: Processes book lending
- `registerBook(bookData)`: Registers new books

### Frontend Architecture
Each HTML file corresponds to a specific functionality:
- `menu.html`: Main navigation menu with links to all features
- `lending.html`: Book checkout interface
- `returning.html`: Book return interface
- `rental_books_finder.html`: Search interface for rented books
- `user_returns.html`: User-specific return interface
- `book_register.html`: New book registration interface

Design characteristics:
- Mobile-first design with large UI elements (48px base font, 72px headings)
- Uses QuaggaJS for barcode scanning functionality
- Inline CSS styling with Google Fonts (Noto Sans JP)
- Client-side JavaScript communicates with backend via `google.script.run`

## Important Notes
- The application has public access (`ANYONE_ANONYMOUS`) configured in appsscript.json
- Email functionality is integrated for sending notifications
- The system uses Stackdriver for exception logging
- All text and comments are in Japanese