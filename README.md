# DataImportLab

A hands-on ASP.NET Core project for practicing customer data imports before applying the same concepts to a real application.

The planned workflow is **Excel → staging → validation → customers**, followed by XML inserts and updates with duplicate-operation protection. Development proceeds in small, working steps, starting with synchronous request processing.

**Current milestone:** upload an Excel workbook, create an import job, validate its worksheet and headers, and write raw customer rows to SQL Server staging in batches. Customer-value validation, promotion to `customers`, and XML processing are still planned.

## Current capabilities

- Accept `.xlsx` uploads through a controller endpoint using `multipart/form-data`.
- Locate the `Customers` worksheet, regardless of its position in the workbook.
- Validate the expected column count and header order.
- Read rows using ExcelDataReader and expose them as `IEnumerable<RawCustomerRow>`.
- Preserve the source row number for future validation error reporting.
- Represent incoming fields as nullable text, using invariant formatting for numbers and `yyyy-MM-dd` for parsed dates.
- Create a `RUNNING` import job and write staging rows using `SqlBulkCopy` in batches of 1,000.
- Commit all staging batches and the job's total count in one transaction, then leave the job `PENDING` for future validation.
- Roll back staging on failure and attempt to record `FAILED` on the separately persisted job.
- Return the filename, file size, job ID, number of staged rows, and job status.
- Return HTTP 400 for known worksheet/header errors through a feature-specific exception.
- Expose OpenAPI and Scalar in the Development environment.
- Register feature and persistence dependencies through `IServiceCollection` extension methods.
- Enforce file-scoped namespaces with compiler-build and editor warnings.

The endpoint stages all rows returned after the header, including rows containing invalid business data. It does not yet save the original file, validate customer values, or insert accepted customers into `customers`.

## Technology

| Component | Package / framework | Version in this repository |
| --- | --- | --- |
| API | ASP.NET Core / .NET | 10 |
| Excel reader | `ExcelDataReader` | 3.9.0 |
| OpenAPI generation | `Microsoft.AspNetCore.OpenApi` | 10.0.11 |
| Interactive API documentation | `Scalar.AspNetCore` | 2.17.2 |
| SQL Server client | `Microsoft.Data.SqlClient` | 7.0.2 |
| Local database target | SQL Server Express | Local named instance |

The project file is the source of truth for package versions: [DataImportLab.Api.csproj](DataImportLab.Api/DataImportLab.Api.csproj).

## Architecture

The application uses feature folders with a vertical-slice direction inside a single API project. HTTP handling and Excel parsing have separate responsibilities. Shared database connection infrastructure lives under `Infrastructure/Persistence`.

```text
DataImportLab/
├── .editorconfig
├── DataImportLab.slnx
├── README.md
└── DataImportLab.Api/
    ├── Features/
    │   └── CustomerImports/
    │       ├── CustomerImportsServiceCollection.cs
    │       └── ImportExcel/
    │           ├── ImportExcelController.cs
    │           ├── ImportExcelHandler.cs
    │           ├── ImportExcelJobRepository.cs
    │           ├── ImportExcelStagingRepository.cs
    │           ├── CustomerExcelReader.cs
    │           ├── InvalidExcelImportException.cs
    │           └── Models/
    │               ├── ImportExcelResult.cs
    │               └── RawCustomerRow.cs
    ├── Infrastructure/
    │   └── Persistence/
    │       ├── PersistenceServiceCollection.cs
    │       └── SqlConnectionFactory.cs
    ├── Properties/
    │   └── launchSettings.json
    ├── appsettings.json
    ├── appsettings.Development.json
    └── Program.cs
```

`ImportXml` is reserved as a future feature; there is no XML implementation yet.

The current request flow is:

```text
Multipart upload
    → ImportExcelController
    → ImportExcelHandler creates a RUNNING job
    → transaction: read and check workbook → stage batches → update job
    → commit → HTTP response with jobId, totalRows, and PENDING status

Failure after job creation
    → rollback staging transaction → mark job FAILED → propagate error
```

- **Controller:** owns the uploaded stream, invokes the handler with a filename and cancellation token, and maps expected import failures to HTTP responses.
- **Handler:** coordinates job creation, reading, the staging transaction, and job status updates. It does not return HTTP responses.
- **Job repository:** creates job records and updates their status and counters with parameterized SQL.
- **Staging repository:** maps raw rows to a reusable `DataTable` of at most 1,000 rows and writes each batch through `SqlBulkCopy` with explicit column mappings.
- **Reader:** knows the Excel layout and returns raw rows. It has no dependency on HTTP responses or SQL Server.
- **RawCustomerRow:** a record with `SourceRow` and nullable string properties for incoming values. It represents unvalidated data.
- **SqlConnectionFactory:** creates a new, closed `SqlConnection` on each call. The caller will open and dispose the connection. The singleton registration applies to the factory, not to a shared connection.
- **Registration extensions:** `AddCustomerImports()` and `AddPersistence(configuration)` keep individual service registrations out of `Program.cs`.

The reader uses deferred execution: reading and exceptions occur while the staging repository enumerates the sequence. The controller awaits the handler inside its `try` block while the uploaded stream remains open. The reader leaves that stream open for its owner to dispose.

This avoids building an application-level list of all customer rows. It does not imply that HTTP uploads are unbuffered: `IFormFile` uses ASP.NET Core's upload handling, which can buffer content in memory or temporary files.

## Local setup

### Prerequisites

- The [.NET 10 SDK](https://dotnet.microsoft.com/en-us/download/dotnet/10.0).
- A development environment supporting .NET 10, or a terminal.
- An `.xlsx` workbook matching the format below. Microsoft Excel does not need to be installed on the API host.
- A local SQL Server Express instance and a database named `ImportDb`, with the practice SQL Server schema applied.

The documented SQL configuration uses Windows Authentication. Application startup requires a nonempty `ImportDb` connection string. Upload requests now require a reachable database with the expected tables and write permissions for the Windows identity running the API.

### Configure the connection string

Add the following section to `DataImportLab.Api/appsettings.Development.json`, preserving any existing settings:

```json
{
  "ConnectionStrings": {
    "ImportDb": "Server=localhost\\SQLEXPRESS;Database=ImportDb;Integrated Security=True;Encrypt=True;TrustServerCertificate=True"
  }
}
```

- `localhost\\SQLEXPRESS` is the JSON representation of the local `localhost\SQLEXPRESS` instance name. Adjust it if your instance differs.
- `Integrated Security=True` uses the Windows identity running the API.
- `Encrypt=True;TrustServerCertificate=True` encrypts the connection while bypassing certificate validation for this local development setup. Use a trusted certificate and certificate validation when configuring a deployed environment.

For other environments, supply `ConnectionStrings:ImportDb` through configuration, such as the `ConnectionStrings__ImportDb` environment variable. Keep credentials, if used in a different setup, outside committed configuration files.

If startup reports `Λείπει το connection string ImportDb.`, check the configuration key as well as its value. `AddPersistence` looks for **`ImportDb`** specifically. This configuration key is separate from the database name, **`ImportDb`**, inside the connection string.

### Restore and build

Before uploading a workbook, create `ImportDb` and apply the separate practice pack's `schema_mssql.sql` to it. In SSMS, select `ImportDb` as the query database before executing the schema. If using `sqlcmd`, pass `-d ImportDb -I`; `-I` enables `QUOTED_IDENTIFIER`, which is required by the computed-column index in the schema.

Run from the repository root:

```shell
dotnet restore DataImportLab.slnx
dotnet build DataImportLab.slnx --no-restore
```

### Run with HTTPS

Trust the local development certificate if needed:

```shell
dotnet dev-certs https --trust
```

Start the API:

```shell
dotnet run --project DataImportLab.Api --launch-profile https
```

Open [Scalar](https://localhost:7250/scalar) in your browser. The OpenAPI document is available at [openapi/v1.json](https://localhost:7250/openapi/v1.json).

The HTTPS launch profile also binds HTTP on port `5085`; the application uses HTTPS redirection.

### Visual Studio

Open `DataImportLab.slnx`, select the API's `https` profile, and run with **F5** or **Ctrl+F5**. Both launch profiles have browser launch enabled and point to `scalar`. When using `dotnet run`, open the browser manually.

An HTTP-only development profile is also available:

```shell
dotnet run --project DataImportLab.Api --launch-profile http
```

For that profile, open [http://localhost:5085/scalar](http://localhost:5085/scalar).

## Excel input contract

Upload an `.xlsx` workbook with a worksheet named **`Customers`**. Worksheet-name matching ignores case. The first row must contain these nine headers in this exact order:

| Column | Header | Raw model property |
| --- | --- | --- |
| A | `customer_id` | `CustomerId` |
| B | `first_name` | `FirstName` |
| C | `last_name` | `LastName` |
| D | `email` | `Email` |
| E | `vat_number` | `VatNumber` |
| F | `status` | `Status` |
| G | `balance` | `Balance` |
| H | `created_at` | `CreatedAt` |
| I | `updated_at` | `UpdatedAt` |

Header comparisons ignore case and trim surrounding whitespace. Reordered, missing, unexpected, or additional columns are rejected. The column-count check uses ExcelDataReader's `FieldCount` for the worksheet.

Data starts on Excel row **2**. `SourceRow` retains the Excel row number rather than a zero-based index.

Example data row:

```text
C000001 | Alex | Example | alex@example.com | 012345678 | ACTIVE | 125.50 | 2025-01-01 | 2025-01-02
```

Store identifiers such as `vat_number` as text when leading zeros matter. A number displayed with a custom Excel format is still a numeric value; the current reader does not reconstruct display formatting.

Existing string values are preserved, including invalid values. Parsed numbers use invariant culture and parsed `DateTime` values become `yyyy-MM-dd`. Cell formatting and time-of-day information are not retained. Business validation is a later step.

A worksheet with no readable rows is rejected. A worksheet containing only valid headers creates a pending job with `totalRows: 0`. There is no explicit blank-data-row filter. Values exceeding the staging capacity (100 characters for `customer_id`, 4,000 for other raw fields) reject the file without truncation; all its staging batches are rolled back.

## API usage

### Upload and stage customer rows

```http
POST /api/imports/customers/excel
Content-Type: multipart/form-data
```

| Form field | Type | Required |
| --- | --- | --- |
| `file` | Uploaded `.xlsx` file | Yes |

In Scalar, select the endpoint, choose the file in the multipart request form, and send the request. Upload the file contents; supplying a filename or local path as a text field does not upload a file.

Example using `curl.exe` on Windows after trusting the development certificate:

```shell
curl.exe --request POST "https://localhost:7250/api/imports/customers/excel" --form "file=@C:/path/to/customers.xlsx"
```

Illustrative successful response:

```json
{
  "fileName": "customers.xlsx",
  "sizeInBytes": 1074829,
  "jobId": 1,
  "totalRows": 20000,
  "status": "PENDING"
}
```

`sizeInBytes` depends on the uploaded file, and `jobId` is generated by SQL Server. `totalRows` excludes the header and counts staged rows, not validated customers. Re-uploading a workbook creates a separate job and a separate set of staging rows.

`PENDING` means staging succeeded and the job is awaiting the future validation step. There is no background worker or automatic continuation yet. `processed_rows`, `accepted_rows`, and `rejected_rows` remain zero, and `completed_at` remains null. A successful staging response does not mean the full customer import is complete.

### Error responses

The reader throws `InvalidExcelImportException` for known worksheet/header failures. The controller maps it to HTTP **400** with a Problem Details response:

```json
{
  "type": "https://tools.ietf.org/html/rfc9110#section-15.5.1",
  "title": "Μη έγκυρο αρχείο Excel",
  "status": 400,
  "detail": "Στη στήλη 4 αναμένεται το header email, αλλά βρέθηκε email_to_send.",
  "traceId": "<request-specific trace identifier>"
}
```

Current application error messages are in Greek.

| Condition | Current behavior |
| --- | --- |
| Empty uploaded file | HTTP 400 with `Το αρχείο είναι κενό.`; this direct `BadRequest` response is not the custom Problem Details response above |
| Missing `Customers` worksheet | HTTP 400, Problem Details |
| Worksheet with no readable rows | HTTP 400, Problem Details |
| Unexpected column count or header | HTTP 400, Problem Details |
| Filename outside 1–255 characters, or raw value exceeding staging capacity | HTTP 400, Problem Details |
| Invalid customer values that fit staging | Stored as raw text; business validation is not implemented |
| Corrupt, unsupported, or password-protected workbook | Library exceptions are not yet translated into the custom HTTP 400 response and can result in HTTP 500 |

After a job is created, failures roll back the staging transaction. The handler then attempts to mark the job `FAILED` with a completion timestamp, using a separate connection and a 10-second timeout even if the request was cancelled. If SQL Server is unavailable for that update, the failure is logged and the job can remain `RUNNING`. The original exception is preserved. Import error-table persistence is not implemented yet.

## Database direction

The local practice database is named `ImportDb`. The applied SQL Server schema consists of:

| Table | Intended responsibility |
| --- | --- |
| `import_jobs` | Import lifecycle, source file, timestamps, and counters |
| `staging_customers` | Incoming text values, linked to a job and source row |
| `import_errors` | Validation failures and their source locations |
| `customers` | Accepted customer records |
| `processed_operations` | XML operation IDs already applied |

The SQL scripts and synthetic practice files are currently maintained outside this repository. There is no migration or automatic schema creation in the API. When using the separate practice pack, select **`schema_mssql.sql`**; the original PostgreSQL `schema.sql` is not the selected schema.

Job creation and transactional staging are implemented. The handler creates the audit job outside the staging transaction, so a failed import can retain a job record while its partial staging writes are rolled back. Business validation and transfer to `customers` are the next steps.

## Practice data and expected outcomes

The separate synthetic practice pack contains `customers.xlsx`, `customer_updates.xml`, `schema_mssql.sql`, and `expected_issues.csv`. These files are **not included in this repository**. You can exercise the current endpoint with a workbook you create using the documented input contract.

The Excel practice file contains 20,000 data rows, including 10 intentional problems. The current endpoint should stage and report **20,000**. Once business validation is implemented, the intended result is **19,990 accepted** and **10 rejected**, keeping the earlier occurrence of a duplicate customer ID.

Planned customer checks include required values, field lengths, email format and uniqueness, allowed statuses (`ACTIVE`, `INACTIVE`, `PENDING`), nonnegative balances, nine-digit VAT numbers, and `updated_at >= created_at`. The VAT rule checks digit format, not the Greek tax-number checksum.

The planned XML workflow uses `operationId` to prevent applying the same change twice. Its practice fixture targets 2,000 updates, 450 inserts, and 50 rejected changes.

**Known fixture inconsistency:** the XML inserts currently omit `first_name`, `last_name`, `vat_number`, and `created_at`, which the SQL Server schema requires. The fixture or an explicitly agreed mapping must be corrected before that milestone. The advertised final total of 20,440 customers is a target, not a currently executable or verified result. Missing values must not be silently invented.

## Manual verification

There is currently no automated test project. Useful checks through Scalar are:

| Test | Expected result |
| --- | --- |
| Upload the original practice workbook | HTTP 200, `totalRows: 20000`, `status: PENDING`, and 20,000 staging rows linked to the returned job ID |
| Rename the worksheet to `Other` in a copy | HTTP 400: worksheet not found |
| Rename `email` to `email_to_send` in a copy | HTTP 400 identifying column 4 and both header names |
| Move `Customers` behind another worksheet | Same count; the reader locates the target worksheet first |
| Change header casing or add surrounding whitespace | Accepted |
| Upload a workbook with valid headers and no data rows | HTTP 200, `totalRows: 0` |
| Upload 1,001 validly structured rows | Both the full batch and the final single row are staged |
| Exceed the staging field capacity after the first 1,000 rows | HTTP 400; job becomes `FAILED` and no staging rows remain for it |

Use copies for negative tests so the original fixture retains its intentional business-data errors.

Local smoke checks verified the 20,000-row fixture, partial final batches, preservation of invalid raw values, header failures, and rollback after an already-written batch. The temporary test jobs and staging rows were removed afterward. These checks are not yet a committed automated test suite.

## Coding conventions

Use file-scoped namespaces:

```csharp
namespace DataImportLab.Api.Features.CustomerImports.ImportExcel;
```

The root `.editorconfig` sets `csharp_style_namespace_declarations = file_scoped:warning` and diagnostic `IDE0161` to warning. The API project enables `EnforceCodeStyleInBuild`, so block-scoped namespaces produce warnings during builds as well as in the editor.

Keep feature-specific models and behavior beside their feature, and register dependencies through the corresponding service-collection extension. Shared infrastructure belongs under `Infrastructure` when multiple use cases need it.

## Next milestones

- [x] Add `ImportExcelHandler` to coordinate reading and persistence.
- [x] Create and update `import_jobs` records for staging and failure.
- [x] Load staging rows in batches with `SqlBulkCopy`.
- [ ] Validate incoming values and record errors with source-row references.
- [ ] Insert accepted customers with defined transaction and failure behavior.
- [ ] Add job status and error-retrieval endpoints.
- [ ] Correct and integrate the XML fixture, then implement inserts and updates.
- [ ] Prevent duplicate XML operations using `operationId`.
- [ ] Translate expected file-parser failures into consistent client errors.
- [ ] Add automated coverage for parsing, validation, and persistence behavior.
- [ ] Explore background processing, progress reporting, and retries after the synchronous flow works.

## License

This repository does not currently include a project license. Dependency licenses remain separate; ExcelDataReader is distributed under the [MIT license](https://github.com/ExcelDataReader/ExcelDataReader/blob/develop/LICENSE).
