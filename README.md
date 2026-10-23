# Compensation Data Reconciliation & Document Automation

**From multiple Excel files to structured records and personalized Word documents.**

I developed a spreadsheet-based workflow to consolidate compensation records, reconcile land information, standardize asset details, and prepare data for document generation at a Commune People’s Committee in Vietnam.

Across the original assignment, I worked with **approximately 1,100 compensation records spanning nearly 11,500 Excel rows**. This repository presents a smaller sample of that workflow, including source workbooks, a consolidated workbook, a Word Mail Merge template, and two document-cleanup macros.

**Author:** Nguyễn Đức Việt  
**Role:** Data Processing & Quality Control  
**Period:** August–September 2026  
**Tools:** Excel · FILTER · XLOOKUP · Conditional Formatting · Group/Outline · Keyboard Macros · Word Mail Merge · VBA

[Explore the consolidated workbook](Tong%20hop%20thon%20V%20sample.xlsx) · [View the source files](raw%20data/) · [Download the Word template](mailling%20and%20vba.doc)

## The operational problem

Several contributors entered compensation information in separate Excel files. Each file combined landholder information, parcel identifiers, asset details, and calculations in a layout designed for human review. A record could contain several structures and crops, while an asset record could relate to more than one parcel or landholder.

This created three recurring challenges:

- **Consistency:** record numbers, land areas, and asset descriptions needed to be checked against reference information.
- **Traceability:** records had to remain connected to their source files and supporting asset details after consolidation.
- **Document preparation:** variable-length asset sections needed to become a fixed set of fields that Word could merge into individual documents.

My task was to connect these stages into a practical preparation and review workflow.

## Workflow

```mermaid
flowchart LR
    A["Contributor workbooks<br/>raw data/"] --> B["Consolidate and organize<br/>File tổng"]
    B --> C["Reconcile record numbers<br/>and land areas"]
    R["Reference data<br/>check thong tin ho"] --> C
    C --> D["Review discrepancies<br/>and correct accepted values"]
    D --> E["Standardize asset details<br/>and prepare calculations"]
    E --> F["Reshape to one row<br/>per document: mailing"]
    F --> G["Word Mail Merge<br/>Asset verification template"]
    G --> H["VBA cleanup<br/>Round decimals and remove blank rows"]
    H --> I["Individual documents<br/>for final review"]

    classDef input fill:#eef5ff,stroke:#5374a8,color:#183047;
    classDef process fill:#edf7f2,stroke:#418466,color:#173b2d;
    classDef output fill:#fff4dc,stroke:#b58b34,color:#493710;
    class A,R input;
    class B,C,D,E,F process;
    class G,H,I output;
```

The arrows describe the working process. Consolidation, exception review, and preparation of the `mailing` sheet include manual steps; the repository does not implement an automatic refresh across every stage.

## My contribution

### Consolidation and source tracking

I combined the contributor files into a master workbook, retained numbered source-file separators, and used Excel’s **Group/Outline** feature to navigate between source sections and their detailed records.

The sample preserves this organization in sheet `File tổng`. The source markers make it easier to trace a section back to its input file, while grouping keeps a long worksheet manageable during review.

### Reference-based reconciliation

I separated the main land records from detail rows, then compared three fields with reference information: the record number (`STT`), total agricultural land area, and acquired land area.

The `check thong tin ho` sheet makes the comparison visible side by side:

| Range | Purpose |
| --- | --- |
| `A:G` | Reference records, including map-sheet and parcel identifiers |
| `J:P` | Working records extracted from `File tổng` |
| `Q` | Combined comparison result for the three checked fields |
| `R:V` | Reference values returned by XLOOKUP |

`FILTER` extracts rows with nonblank record-number and landholder fields. **XLOOKUP** matches the map-sheet and parcel-number combination, using a separator between the two identifiers. Comparison formulas and conditional formatting direct attention to discrepancies that need review.

Examples from the workbook, shown with standard Excel function names:

```excel
=XLOOKUP(O2&"]"&P2,$F$2:$F$30&"]"&$G$2:$G$30,$A$2:$E$30,,0)

=AND(J2=R2,M2=U2,N2=V2)
```

The comparison result checks those three numeric fields. It does not establish that ownership details or every other field have been verified. Conflicting values require review of the supporting information before correction.

### Asset standardization and repetitive-task automation

I reviewed inconsistent asset names and aligned the names, specifications, and associated values with the project’s reference information. Similar names alone were insufficient: the characteristics of an asset also mattered when determining the appropriate category.

I built **keyboard macros** to assist repeated editing and formula entry across sections with different numbers of asset rows. These supported tasks such as subtotal preparation and separate entries for assets requiring density-related treatment.

The keyboard macros were used in the original workflow. The code supplied in this repository consists of the two **Word VBA** procedures described below.

### Reshaping records for document generation

I prepared a wide `mailing` table with **one row per document**, pairing each asset name with its specification-and-quantity text. This required preserving the relationship between land records and asset records, including cases involving multiple parcels.

In this sample, **eight main land-record rows are represented in six document rows**. Two document rows combine information associated with more than one parcel. These are different units of observation; worksheet rows, parcels, asset records, and documents should not be counted interchangeably.

The broader assignment included compensation-notice preparation. The Word template supplied here demonstrates the same data-to-document approach through an **asset verification record**: *Biên bản làm việc về việc xác minh nguồn gốc, thời điểm hình thành tài sản*.

## Repository contents

```text
.
├── README.md
├── raw data/
│   ├── Thôn V (tập 1).xlsx
│   ├── Thôn V (tập 2).xlsx
│   └── Thôn V (tập 3).xlsx
├── Tong hop thon V sample.xlsx
├── mailling and vba.doc
├── vba 1 remove digits.txt
└── vba 2 remove totally blank rows in table.txt
```

| Artifact | What to inspect |
| --- | --- |
| [Source workbooks](raw%20data/) | Mixed record/detail layouts, asset calculations, and supporting crop and structure reference sheets |
| [Consolidated workbook](Tong%20hop%20thon%20V%20sample.xlsx) | `File tổng`, `check thong tin ho`, and `mailing` |
| [Word template](mailling%20and%20vba.doc) | Merge fields for land information and variable-length asset lists |
| [Decimal-formatting macro](vba%201%20remove%20digits.txt) | `LamTronSoThapPhan` |
| [Blank-row cleanup macro](vba%202%20remove%20totally%20blank%20rows%20in%20table.txt) | `XoaDongTrangHoanToan` and its helper, `NoiDungO` |

### Sample structure

| Component | Included sample |
| --- | ---: |
| Contributor workbooks | 3 |
| Main land-record rows in `File tổng` | 8 |
| Reference records in `check thong tin ho` | 29 |
| Document rows in `mailing` | 6 |
| Columns in `mailing` | 34 |
| Structure slots per document | 7 |
| Crop/livestock slots per document | 5 |
| Word VBA source files | 2 |

These counts describe the repository sample. The approximately 1,100 records and 11,500 rows refer to the complete original assignment.

## Mail Merge data structure

The `mailing` sheet keeps the original field names used for document preparation.

| Field family | Meaning |
| --- | --- |
| `STT` | Record or asset-record identifier |
| `chu dat`, `chu tai san` | Landholder and asset-owner information |
| `tong dien tich`, `dtth` | Total agricultural land area and acquired area |
| `to`, `thua` | Map-sheet and parcel information; a document may list multiple parcels |
| `TS1`–`TS7` | Structure and construction-item names |
| `SLTS1`–`SLTS7` | Corresponding specification-and-quantity text |
| `CT1`–`CT5` | Crop/livestock item names |
| `SLCT1`–`SLCT5` | Corresponding specification-and-quantity text |
| Final three amount columns | Structure amount, crop amount, and total asset compensation |

The `SLTS` and `SLCT` fields contain descriptive text, not just numeric quantities. Empty asset positions remain blank so the same Word template can serve records with fewer items.

The Word template uses names such as `chu_dat` and `chu_tai_san`. When reconnecting it, check their mapping to the Excel headers with spaces. The supplied template uses 30 merge fields; it does not display every column in `mailing`, including the final compensation totals.

## Word VBA cleanup

### Decimal formatting

`LamTronSoThapPhan` searches the main document body for numbers matching its decimal pattern and formats them to **one decimal place**. Despite the filename “remove digits,” the procedure performs numeric formatting rather than simply deleting characters.

Its conversion uses `CDbl`, so the decimal separator and regional settings matter. The template already applies a one-decimal numeric format to the `dtth` merge field; the macro also addresses matching numbers elsewhere in the document body.

### Removing unused table rows

`XoaDongTrangHoanToan` traverses the document’s tables from the bottom upward and deletes a row only when **all its cells are empty** after cleanup. It preserves the first row of each table.

The helper `NoiDungO` removes Word’s cell/end-of-paragraph markers and selected whitespace characters before checking whether a cell contains text. This allows unused asset slots to be removed while retaining rows that contain asset information or section labels.

Run document cleanup on the **merged output**, where unused fields have become empty. The unmerged template still contains merge fields and should remain available for reuse.

## How to explore the project

1. **Review the inputs.** Open the workbooks in `raw data/` and locate the main records and their asset sections. The reference sheets provide context for the original layout.
2. **Follow consolidation.** Open `Tong hop thon V sample.xlsx`, then inspect the source separators and grouped sections in `File tổng`.
3. **Inspect the checks.** In `check thong tin ho`, compare the working values with the XLOOKUP results. Investigate `FALSE` results before treating a record as reviewed.
4. **Inspect the output shape.** In `mailing`, follow how the asset lists become paired name/detail fields. Review the source-to-document mapping, particularly for multi-parcel records.
5. **Reconnect Word.** Open a working copy of `mailling and vba.doc`, select the consolidated workbook as its Mail Merge data source, and select the `mailing` worksheet. Review field mappings and preview individual records.
6. **Merge and clean up.** Generate a new merged document. To test the VBA, add the supplied procedures to a standard module in Word, including `NoiDungO` with the blank-row procedure, then run them on the output copy.
7. **Review the result.** Check the record identifiers, parcel information, asset descriptions, quantities, and remaining table rows against the reviewed source data.

Use Excel with support for `FILTER` and `XLOOKUP`, and desktop Word for VBA. Download the Office files to inspect their full behavior; the VBA `.txt` files can also be read directly on GitHub.

## Current sample status and limitations

The files are working snapshots of the process. The current sample retains **three `FALSE` reconciliation results**, and one missing land-area input produces dependent `#DIV/0!` results in `File tổng`. The source workbooks also retain incomplete calculations and broken references.

The `mailing` table is largely a prepared snapshot, so changes to `File tổng` do not automatically update all document fields. Its values need a final reconciliation with the accepted source records before generating documents. The supplied workbooks also retain legacy external-workbook links; those linked files are not dependencies supplied by this repository.

For a repeatable version, I would add explicit checks for duplicate lookup keys, a review log recording accepted corrections, a refreshable transformation into the merge table, and a capacity check for records exceeding seven structure or five crop slots. The current XLOOKUP uses the first exact match; duplicate-key validation is a future improvement.

The sample demonstrates data-processing methods, not guidance on compensation policy or entitlement. No percentage improvement in accuracy or processing time is claimed, because those outcomes were not formally measured.

## What this project demonstrates

- **Data preparation:** organizing irregular, multi-file inputs while retaining source context.
- **Data quality:** defining field-level checks and making discrepancies visible for review.
- **Data modeling:** separating parcel-level information, asset details, and document-level output.
- **Practical automation:** combining spreadsheet formulas, keyboard macros, Mail Merge, and Word VBA around an operational task.
- **Judgment:** recognizing where a formula can identify an issue and where a person must decide how to resolve it.

The result is a documented workflow that connects spreadsheet preparation with the documents staff need to produce and review.
