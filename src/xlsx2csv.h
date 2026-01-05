#ifndef _XLSX2CSV_H
#define _XLSX2CSV_H

#include <stdbool.h>
#include <stdint.h>
#include <stdio.h>
#include <time.h>

/* Version information */
#define XLSX2CSV_VERSION "0.8.3"

/* CSV quoting modes (compatible with Python csv module) */
typedef enum {
    QUOTE_MINIMAL    = 0,
    QUOTE_ALL        = 1,
    QUOTE_NONNUMERIC = 2,
    QUOTE_NONE       = 3
} quotingMode;

/* Format types */
typedef enum {
    FORMAT_STRING,
    FORMAT_FLOAT,
    FORMAT_CUSTOM_FLOAT, /* Custom float format (e.g. "0_ ") - requires data validation */
    FORMAT_DATE,
    FORMAT_TIME,
    FORMAT_BOOLEAN,
    FORMAT_PERCENTAGE
} formatType;

/* Options structure */
typedef struct {
    char        delimiter;
    quotingMode quoting;
    char       *sheetdelimiter;
    char       *dateformat;
    char       *timeformat;
    char       *floatformat;
    bool        scifloat;
    bool        skip_empty_lines;
    bool        skip_trailing_columns;
    bool        escape_strings;
    bool        no_line_breaks;
    bool        hyperlinks;
    char      **include_sheet_pattern;
    int         include_sheet_pattern_count;
    char      **exclude_sheet_pattern;
    int         exclude_sheet_pattern_count;
    bool        exclude_hidden_sheets;
    bool        merge_cells;
    char       *outputencoding;
    char       *lineterminator;
    char      **ignore_formats;
    int         ignore_formats_count;
    bool        skip_hidden_rows;
} xlsxOptions;

/* Sheet information */
typedef struct {
    char *name;
    char *relation_id;
    int   index;
    char *state;
} sheetInfo;

/* Workbook structure */
typedef struct {
    sheetInfo *sheets;
    int        sheet_count;
    bool       date1904;
} workbookInfo;

/* Shared strings */
typedef struct {
    char **strings;
    int    count;
} sharedStrings;

/* Number format */
typedef struct {
    int        id;
    char      *format_code;
    formatType type;
} numFormat;

/* Style information */
typedef struct {
    numFormat *formats;
    int        format_count;
    int       *cell_xfs;
    int        cell_xfs_count;
} styleInfo;

/* Main converter structure */
typedef struct {
    void         *zip_handle;
    xlsxOptions   options;
    workbookInfo  workbook;
    sharedStrings shared_strings;
    styleInfo     styles;
    bool          has_date_error; /* Flag for date format errors (Python compatibility) */
} xlsx2csvConverter;

/* Main API functions */
xlsx2csvConverter *xlsx2csv_create(const char *filename, xlsxOptions *options);
void               xlsx2csv_free(xlsx2csvConverter *conv);
int                xlsx2csv_convert(xlsx2csvConverter *conv,
                                    const char        *outfile,
                                    int                sheetid,
                                    const char        *sheetname);
int                xlsx2csv_convert_all(xlsx2csvConverter *conv, const char *outdir);

#endif /* _XLSX2CSV_H */
