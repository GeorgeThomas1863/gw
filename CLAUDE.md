# Access VBA SharePoint Database Project

## Project Overview

This is a COPY of a Microsoft Access application that provides a user interface for managing data stored in SharePoint. The application uses VBA forms to accept user input, search/query data, and manipulate records stored in specific SharePoint lists.

## Project Structure

```
/forms/          - Form class modules (.cls files) - UI and event handlers
/modules/        - Standard VBA modules (.bas files) - Business logic and utilities
```

## Key Components

### Forms (`/forms/*.cls`)

Form class modules handle user interface and events. Each form typically contains:

- Event handlers (button clicks, form load/unload, etc.)
- Input validation logic
- Calls to business logic in standard modules
- UI state management

### Modules (`/modules/*.bas`)

Standard modules contain reusable code:

- SharePoint connection and data access functions
- Data validation and transformation logic
- Utility functions used across multiple forms
- Constants and configuration

## Technology Stack

- **VBA (Visual Basic for Applications)** - Primary language
- **Microsoft Access** - Host application and local UI
- **SharePoint** - Backend data storage (lists/libraries)
- Likely uses SharePoint REST API or SOAP services for data operations

## Development Workflow

**IMPORTANT**: THIS IS A COPY OF A SEPARATE Access database file. Changes made here to `.cls` and `.bas` files must be imported back into the Access database file (.accdb) to take effect. The user manually created a copy of this project so you can see it here, but the actual runtime is the Access database. The files here have NO connection to / integration with the actual Access database and will be imported back to the Access database by the user at a later time.

### Making Changes

1. Edit `.cls` or `.bas` files in this IDE
2. The user will import and test all changes in the actual Access runtime environment.

## Common Tasks

### SharePoint Integration Patterns

Look for these common patterns in the code:

- HTTP requests to SharePoint REST endpoints
- XML/JSON parsing for SharePoint responses
- Authentication (could be Windows Auth, OAuth, or SharePoint App credentials)
- CRUD operations (Create, Read, Update, Delete) on SharePoint lists

### Code Review & Refactoring

When reviewing or refactoring:

- Check for error handling (`On Error GoTo` patterns)
- Verify proper resource cleanup (closing connections, clearing objects)
- Ensure consistent naming conventions across forms and modules

## Things to Know

### VBA Limitations

- No native JSON parsing (uses custom json parser)
- Limited async capabilities - most operations are synchronous
- String manipulation can be verbose
- COM object management requires explicit cleanup

### Access-Specific Considerations

- Forms have a lifecycle (Load, Current, Unload events)
- RecordSource property may bind forms to local tables/queries
- DoCmd object used for Access-specific actions (navigation, etc.)
- References may include: Microsoft ActiveX Data Objects, XML libraries, etc.

### Testing

Changes must be tested in the actual Access application. The text files alone won't run - they need to be in an Access database context. The user will handle all testing.

## File Naming Conventions

- Form files: Typically named `Form_[FormName].cls` or `[FormName].cls`
- Module files: Descriptive names like `SharePointUtils.bas`, `DataValidation.bas`, etc.

## Getting Started

To understand this codebase:

1. Start with form modules to understand the user workflows
2. Trace button click events to see what business logic they call
3. Review modules to understand SharePoint integration patterns
4. Look for a main/entry point or startup form
5. Check for configuration constants (URLs, list names, etc.)

---

**Note to Claude Code**: This is a VBA project exported as text files. While you can edit these files directly, remember that the person will need to import your changes back into their Access database for testing. When suggesting changes, consider the VBA language constraints and Access runtime environment.
