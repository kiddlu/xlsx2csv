#ifndef XLSX2CSV_H
#define XLSX2CSV_H

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
} xlsx2csvConverter;

/* Function declarations */

/* zip_reader.c */
void *zip_open_file(const char *filename);
void *zip_open_stdin(void);
void  xlsx_zip_close(void *handle);
void *zip_file_open(void *zip_handle, const char *filename);
int   zip_file_read(void *file_handle, void *buffer, size_t size);
void  zip_file_close(void *file_handle);
char *zip_read_file_to_string(void *zip_handle, const char *filename);

/* xml_parser.c */
int parse_content_types(xlsx2csvConverter *conv);
int parse_workbook(xlsx2csvConverter *conv);
int parse_shared_strings(xlsx2csvConverter *conv);
int parse_styles(xlsx2csvConverter *conv);
int parse_worksheet(xlsx2csvConverter *conv, int sheet_index, FILE *outfile);

/* format_handler.c */
char      *format_cell_value(const char        *value,
                             const char        *type_attr,
                             const char        *style_attr,
                             xlsx2csvConverter *conv);
formatType get_format_type(int style_id, styleInfo *styles);
char      *format_date(double value, const char *format, bool date1904);
char      *format_time(double value, const char *format);
char      *format_float(double value, const char *format, bool scifloat, const char *original_str);

/* csv_writer.c */
typedef struct csvWriter csvWriter;
csvWriter               *csv_writer_create(FILE *fp, xlsxOptions *options);
void                     csv_writer_free(csvWriter *writer);
int                      csv_write_row(csvWriter *writer, char **fields, int field_count);
int                      csv_write_field(csvWriter *writer, const char *field);
void                     csv_writer_reset_row(csvWriter *writer);
void                     csv_writer_set_field_count(csvWriter *writer, int count);

/* utils.c */
char *str_duplicate(const char *str);
char *str_concat(const char *str1, const char *str2);
bool  str_match_pattern(const char *str, const char *pattern);
int   column_name_to_index(const char *column);
char *column_index_to_name(int index);
bool  is_numeric(const char *str);

/* main.c / xlsx2csv.c */
xlsx2csvConverter *xlsx2csv_create(const char *filename, xlsxOptions *options);
void               xlsx2csv_free(xlsx2csvConverter *conv);
int                xlsx2csv_convert(xlsx2csvConverter *conv,
                                    const char        *outfile,
                                    int                sheetid,
                                    const char        *sheetname);
int                xlsx2csv_convert_all(xlsx2csvConverter *conv, const char *outdir);

#endif /* XLSX2CSV_H */
