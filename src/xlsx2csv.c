/* Standard library headers */
#include <stdio.h>
#include <stdlib.h>
#include <string.h>

/* Project headers */
#include "csv_writer.h"
#include "format_handler.h"
#include "utils.h"
#include "xlsx2csv.h"
#include "xml_parser.h"
#include "zip_reader.h"

/* Initialize options with defaults */
static void init_default_options(xlsxOptions *opts)
{
    opts->delimiter                   = ',';
    opts->quoting                     = QUOTE_MINIMAL;
    opts->sheetdelimiter              = str_duplicate("--------");
    opts->dateformat                  = NULL;
    opts->timeformat                  = NULL;
    opts->floatformat                 = NULL;
    opts->scifloat                    = false;
    opts->skip_empty_lines            = false;
    opts->skip_trailing_columns       = false;
    opts->escape_strings              = false;
    opts->no_line_breaks              = false;
    opts->hyperlinks                  = false;
    opts->include_sheet_pattern       = NULL;
    opts->include_sheet_pattern_count = 0;
    opts->exclude_sheet_pattern       = NULL;
    opts->exclude_sheet_pattern_count = 0;
    opts->exclude_hidden_sheets       = false;
    opts->merge_cells                 = false;
    opts->outputencoding              = str_duplicate("utf-8");
    opts->lineterminator              = str_duplicate("\n");
    opts->ignore_formats              = NULL;
    opts->ignore_formats_count        = 0;
    opts->skip_hidden_rows            = true;
}

/* Create xlsx2csv converter */
xlsx2csvConverter *xlsx2csv_create(const char *filename, xlsxOptions *options)
{
    xlsx2csvConverter *conv = calloc(1, sizeof(xlsx2csvConverter));
    if (!conv) {
        return NULL;
    }

    /* Initialize error flag */
    conv->has_date_error = false;

    /* Initialize options */
    if (options) {
        memcpy(&conv->options, options, sizeof(xlsxOptions));
    } else {
        init_default_options(&conv->options);
    }

    /* Open ZIP file */
    if (filename && strcmp(filename, "-") == 0) {
        conv->zip_handle = zip_open_stdin();
    } else {
        conv->zip_handle = zip_open_file(filename);
    }

    if (!conv->zip_handle) {
        free(conv);
        return NULL;
    }

    /* Parse metadata */
    if (parse_content_types(conv) < 0) {
        fprintf(stderr, "Warning: Failed to parse content types\n");
    }

    if (parse_workbook(conv) < 0) {
        fprintf(stderr, "Error: Failed to parse workbook\n");
        xlsx2csv_free(conv);
        return NULL;
    }

    if (parse_shared_strings(conv) < 0) {
        fprintf(stderr, "Warning: Failed to parse shared strings\n");
    }

    if (parse_styles(conv) < 0) {
        fprintf(stderr, "Warning: Failed to parse styles\n");
    }

    return conv;
}

/* Free converter */
void xlsx2csv_free(xlsx2csvConverter *conv)
{
    if (!conv) {
        return;
    }

    /* Free ZIP handle */
    if (conv->zip_handle) {
        xlsx_zip_close(conv->zip_handle);
    }

    /* Free workbook data */
    for (int i = 0; i < conv->workbook.sheet_count; i++) {
        free(conv->workbook.sheets[i].name);
        free(conv->workbook.sheets[i].relation_id);
        free(conv->workbook.sheets[i].state);
    }
    free(conv->workbook.sheets);

    /* Free shared strings */
    for (int i = 0; i < conv->shared_strings.count; i++) {
        free(conv->shared_strings.strings[i]);
    }
    free(conv->shared_strings.strings);

    /* Free styles */
    for (int i = 0; i < conv->styles.format_count; i++) {
        free(conv->styles.formats[i].format_code);
    }
    free(conv->styles.formats);
    free(conv->styles.cell_xfs);

    /* Note: We don't free option strings as they may point to static strings or command-line
     * arguments */

    free(conv);
}

/* Get sheet index by name */
static int get_sheet_index_by_name(xlsx2csvConverter *conv, const char *sheetname)
{
    for (int i = 0; i < conv->workbook.sheet_count; i++) {
        if (conv->workbook.sheets[i].name &&
            strcmp(conv->workbook.sheets[i].name, sheetname) == 0) {
            return conv->workbook.sheets[i].index;
        }
    }
    return -1;
}

/* Check if sheet should be included based on patterns */
static bool should_include_sheet(const char *sheetname, xlsxOptions *opts)
{
    /* Check include patterns */
    if (opts->include_sheet_pattern_count > 0) {
        bool matched = false;
        for (int i = 0; i < opts->include_sheet_pattern_count; i++) {
            if (str_match_pattern(sheetname, opts->include_sheet_pattern[i])) {
                matched = true;
                break;
            }
        }
        if (!matched) {
            return false;
        }
    }

    /* Check exclude patterns */
    if (opts->exclude_sheet_pattern_count > 0) {
        for (int i = 0; i < opts->exclude_sheet_pattern_count; i++) {
            if (str_match_pattern(sheetname, opts->exclude_sheet_pattern[i])) {
                return false;
            }
        }
    }

    return true;
}

/* Convert single sheet */
int xlsx2csv_convert(xlsx2csvConverter *conv,
                     const char        *outfile,
                     int                sheetid,
                     const char        *sheetname)
{
    if (!conv) {
        return -1;
    }

    /* Resolve sheet name to ID if provided */
    if (sheetname) {
        sheetid = get_sheet_index_by_name(conv, sheetname);
        if (sheetid < 0) {
            fprintf(stderr, "Error: Sheet '%s' not found\n", sheetname);
            return -1;
        }
    }

    /* Open output file */
    FILE *fp = NULL;
    if (!outfile || strcmp(outfile, "-") == 0) {
        fp = stdout;
    } else {
        fp = fopen(outfile, "w");
        if (!fp) {
            fprintf(stderr, "Error: Could not open output file '%s'\n", outfile);
            return -1;
        }
    }

    /* Convert sheet */
    int result = parse_worksheet(conv, sheetid, fp);

    /* Close output file */
    if (fp != stdout) {
        fclose(fp);
    }

    return result;
}

/* Convert all sheets */
int xlsx2csv_convert_all(xlsx2csvConverter *conv, const char *outdir)
{
    if (!conv) {
        return -1;
    }

    for (int i = 0; i < conv->workbook.sheet_count; i++) {
        sheetInfo *sheet = &conv->workbook.sheets[i];

        /* Check if hidden */
        if (conv->options.exclude_hidden_sheets && sheet->state &&
            (strcmp(sheet->state, "hidden") == 0 || strcmp(sheet->state, "veryHidden") == 0)) {
            continue;
        }

        /* Check patterns */
        if (!should_include_sheet(sheet->name, &conv->options)) {
            continue;
        }

        /* Build output filename */
        char outfile[1024];
        snprintf(outfile, sizeof(outfile), "%s/%s.csv", outdir, sheet->name);

        /* Convert sheet */
        printf("Converting sheet '%s' to '%s'\n", sheet->name, outfile);
        if (xlsx2csv_convert(conv, outfile, sheet->index, NULL) < 0) {
            fprintf(stderr, "Error: Failed to convert sheet '%s'\n", sheet->name);
            return -1;
        }
    }

    return 0;
}
