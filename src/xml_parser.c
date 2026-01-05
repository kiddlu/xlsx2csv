/* Standard library headers */
#include <ctype.h>
#include <stdio.h>
#include <stdlib.h>
#include <string.h>

/* Third-party library headers */
#include <expat.h>

/* Project headers */
#include "csv_writer.h"
#include "format_handler.h"
#include "utils.h"
#include "xlsx2csv.h"
#include "xml_parser.h"
#include "zip_reader.h"

/* Count sheets state */
typedef struct {
    int  *sheet_count;
    bool *in_sheets;
} count_sheets_state_t;

static void count_sheets_start(void *userData, const XML_Char *name, const XML_Char **atts)
{
    (void)atts; /* Unused */
    count_sheets_state_t *state = (count_sheets_state_t *)userData;
    if (strcmp(name, "sheets") == 0) {
        *state->in_sheets = true;
    } else if (strcmp(name, "sheet") == 0 && *state->in_sheets) {
        (*state->sheet_count)++;
    }
}

static void count_sheets_end(void *userData, const XML_Char *name)
{
    count_sheets_state_t *state = (count_sheets_state_t *)userData;
    if (strcmp(name, "sheets") == 0) {
        *state->in_sheets = false;
    }
}

/* Count strings state */
typedef struct {
    int  *count;
    bool *in_si;
} count_strings_state_t;

static void count_strings_start(void *userData, const XML_Char *name, const XML_Char **atts)
{
    (void)atts; /* Unused */
    count_strings_state_t *state = (count_strings_state_t *)userData;
    if (strcmp(name, "si") == 0) {
        *state->in_si = true;
        (*state->count)++;
    }
}

static void count_strings_end(void *userData, const XML_Char *name)
{
    count_strings_state_t *state = (count_strings_state_t *)userData;
    if (strcmp(name, "si") == 0) {
        *state->in_si = false;
    }
}

/* Count styles state */
typedef struct {
    int  *format_count;
    int  *xf_count;
    bool *in_num_fmts;
    bool *in_cell_xfs;
} count_styles_state_t;

static void count_styles_start(void *userData, const XML_Char *name, const XML_Char **atts)
{
    (void)atts; /* Unused */
    count_styles_state_t *state = (count_styles_state_t *)userData;
    if (strcmp(name, "numFmts") == 0) {
        *state->in_num_fmts = true;
    } else if (strcmp(name, "numFmt") == 0 && *state->in_num_fmts) {
        (*state->format_count)++;
    } else if (strcmp(name, "cellXfs") == 0) {
        *state->in_cell_xfs = true;
    } else if (strcmp(name, "xf") == 0 && *state->in_cell_xfs) {
        (*state->xf_count)++;
    }
}

static void count_styles_end(void *userData, const XML_Char *name)
{
    count_styles_state_t *state = (count_styles_state_t *)userData;
    if (strcmp(name, "numFmts") == 0) {
        *state->in_num_fmts = false;
    } else if (strcmp(name, "cellXfs") == 0) {
        *state->in_cell_xfs = false;
    }
}

/* Parse Content Types XML */
int parse_content_types(xlsx2csvConverter *conv)
{
    char *xml_data = zip_read_file_to_string(conv->zip_handle, "[Content_Types].xml");
    if (!xml_data) {
        fprintf(stderr, "Error: Could not read [Content_Types].xml\n");
        return -1;
    }

    XML_Parser parser = XML_ParserCreate(NULL);
    if (!parser) {
        free(xml_data);
        return -1;
    }

    int status = XML_Parse(parser, xml_data, (int)strlen(xml_data), 1);
    XML_ParserFree(parser);
    free(xml_data);

    if (!status) {
        fprintf(stderr, "Error: Failed to parse [Content_Types].xml\n");
        return -1;
    }

    /* Content types parsing not critical for basic functionality */
    return 0;
}

/* Workbook parsing state */
typedef struct {
    xlsx2csvConverter *conv;
    int                sheet_count;
    int                current_sheet_idx;
    bool               in_sheets;
    bool               in_sheet;
    bool               in_workbook_pr;
    char              *current_name;
    char              *current_rid;
    char              *current_state;
} workbook_parse_state;

static void workbook_start_element(void *userData, const XML_Char *name, const XML_Char **atts)
{
    workbook_parse_state *state = (workbook_parse_state *)userData;

    if (strcmp(name, "sheets") == 0) {
        state->in_sheets = true;
    } else if (strcmp(name, "sheet") == 0 && state->in_sheets) {
        state->in_sheet = true;
        state->current_sheet_idx++;
        state->current_name  = NULL;
        state->current_rid   = NULL;
        state->current_state = NULL;

        for (int i = 0; atts[i]; i += 2) {
            if (strcmp(atts[i], "name") == 0) {
                state->current_name = str_duplicate(atts[i + 1]);
            } else if (strcmp(atts[i], "r:id") == 0) {
                state->current_rid = str_duplicate(atts[i + 1]);
            } else if (strcmp(atts[i], "state") == 0) {
                state->current_state = str_duplicate(atts[i + 1]);
            }
        }
    } else if (strcmp(name, "workbookPr") == 0) {
        state->in_workbook_pr = true;
        for (int i = 0; atts[i]; i += 2) {
            if (strcmp(atts[i], "date1904") == 0) {
                if (strcmp(atts[i + 1], "true") == 0 || strcmp(atts[i + 1], "1") == 0) {
                    state->conv->workbook.date1904 = true;
                }
            }
        }
    }
}

static void workbook_end_element(void *userData, const XML_Char *name)
{
    workbook_parse_state *state = (workbook_parse_state *)userData;

    if (strcmp(name, "sheets") == 0) {
        state->in_sheets = false;
    } else if (strcmp(name, "sheet") == 0 && state->in_sheet) {
        if (state->current_sheet_idx > 0 && state->current_sheet_idx <= state->sheet_count) {
            int idx                                       = state->current_sheet_idx - 1;
            state->conv->workbook.sheets[idx].name        = state->current_name;
            state->conv->workbook.sheets[idx].relation_id = state->current_rid;
            state->conv->workbook.sheets[idx].state       = state->current_state;
            state->conv->workbook.sheets[idx].index       = idx + 1;
        }
        state->in_sheet = false;
    } else if (strcmp(name, "workbookPr") == 0) {
        state->in_workbook_pr = false;
    }
}

/* Parse Workbook XML */
int parse_workbook(xlsx2csvConverter *conv)
{
    char *xml_data = zip_read_file_to_string(conv->zip_handle, "xl/workbook.xml");
    if (!xml_data) {
        fprintf(stderr, "Error: Could not read xl/workbook.xml\n");
        return -1;
    }

    /* First pass: count sheets */
    XML_Parser parser = XML_ParserCreate(NULL);
    if (!parser) {
        free(xml_data);
        return -1;
    }

    /* Count sheets by parsing again */
    parser = XML_ParserCreate(NULL);
    if (!parser) {
        free(xml_data);
        return -1;
    }

    int                  sheet_count = 0;
    bool                 in_sheets   = false;
    count_sheets_state_t count_state = {&sheet_count, &in_sheets};
    XML_SetUserData(parser, &count_state);
    XML_SetElementHandler(parser, count_sheets_start, count_sheets_end);
    int status = XML_Parse(parser, xml_data, (int)strlen(xml_data), 1);
    XML_ParserFree(parser);

    if (!status) {
        free(xml_data);
        fprintf(stderr, "Error: Failed to parse xl/workbook.xml\n");
        return -1;
    }

    /* Allocate sheet array */
    conv->workbook.sheets      = calloc((size_t)sheet_count, sizeof(sheetInfo));
    conv->workbook.sheet_count = sheet_count;

    /* Second pass: parse sheets */
    parser = XML_ParserCreate(NULL);
    if (!parser) {
        free(xml_data);
        return -1;
    }

    workbook_parse_state parse_state = {0};
    parse_state.conv                 = conv;
    parse_state.sheet_count          = sheet_count;
    XML_SetUserData(parser, &parse_state);
    XML_SetElementHandler(parser, workbook_start_element, workbook_end_element);
    status = XML_Parse(parser, xml_data, (int)strlen(xml_data), 1);
    XML_ParserFree(parser);
    free(xml_data);

    if (!status) {
        fprintf(stderr, "Error: Failed to parse xl/workbook.xml\n");
        return -1;
    }

    return 0;
}

/* Shared strings parsing state */
typedef struct {
    xlsx2csvConverter *conv;
    int                string_idx;
    bool               in_si;
    bool               in_t;
    char              *current_text;
    size_t             text_len;
    size_t             text_capacity;
} shared_strings_state;

static void shared_strings_start_element(void            *userData,
                                         const XML_Char  *name,
                                         const XML_Char **atts)
{
    (void)atts; /* Unused */
    shared_strings_state *state = (shared_strings_state *)userData;

    if (strcmp(name, "si") == 0) {
        state->in_si = true;
        if (state->current_text) {
            free(state->current_text);
            state->current_text = NULL;
        }
        state->text_len      = 0;
        state->text_capacity = 0;
    } else if (strcmp(name, "t") == 0 && state->in_si) {
        state->in_t = true;
    }
}

static void shared_strings_end_element(void *userData, const XML_Char *name)
{
    shared_strings_state *state = (shared_strings_state *)userData;

    if (strcmp(name, "si") == 0) {
        if (state->string_idx < state->conv->shared_strings.count) {
            if (state->current_text) {
                state->conv->shared_strings.strings[state->string_idx] = state->current_text;
                state->current_text                                    = NULL;
            } else {
                state->conv->shared_strings.strings[state->string_idx] = str_duplicate("");
            }
            state->string_idx++;
        }
        state->in_si = false;
    } else if (strcmp(name, "t") == 0) {
        state->in_t = false;
    }
}

static void shared_strings_char_data(void *userData, const XML_Char *s, int len)
{
    shared_strings_state *state = (shared_strings_state *)userData;

    if (state->in_t && state->in_si) {
        size_t new_len = state->text_len + (size_t)len;
        if (new_len >= state->text_capacity) {
            state->text_capacity = new_len + 256;
            state->current_text  = realloc(state->current_text, state->text_capacity + 1);
            if (!state->current_text) {
                return;
            }
        }
        memcpy(state->current_text + state->text_len, s, (size_t)len);
        state->text_len              = new_len;
        state->current_text[new_len] = '\0';
    }
}

/* Parse Shared Strings XML */
int parse_shared_strings(xlsx2csvConverter *conv)
{
    char *xml_data = zip_read_file_to_string(conv->zip_handle, "xl/sharedStrings.xml");
    if (!xml_data) {
        /* No shared strings is valid */
        conv->shared_strings.strings = NULL;
        conv->shared_strings.count   = 0;
        return 0;
    }

    /* First pass: count strings */
    XML_Parser parser = XML_ParserCreate(NULL);
    if (!parser) {
        free(xml_data);
        return -1;
    }

    int                   count       = 0;
    bool                  in_si       = false;
    count_strings_state_t count_state = {&count, &in_si};
    XML_SetUserData(parser, &count_state);
    XML_SetElementHandler(parser, count_strings_start, count_strings_end);
    int status = XML_Parse(parser, xml_data, (int)strlen(xml_data), 1);
    XML_ParserFree(parser);

    if (!status) {
        free(xml_data);
        fprintf(stderr, "Error: Failed to parse xl/sharedStrings.xml\n");
        return -1;
    }

    /* Allocate string array */
    conv->shared_strings.strings = calloc((size_t)count, sizeof(char *));
    conv->shared_strings.count   = count;

    /* Second pass: parse strings */
    parser = XML_ParserCreate(NULL);
    if (!parser) {
        free(xml_data);
        return -1;
    }

    shared_strings_state state = {0};
    state.conv                 = conv;
    XML_SetUserData(parser, &state);
    XML_SetElementHandler(parser, shared_strings_start_element, shared_strings_end_element);
    XML_SetCharacterDataHandler(parser, shared_strings_char_data);
    status = XML_Parse(parser, xml_data, (int)strlen(xml_data), 1);
    XML_ParserFree(parser);
    free(xml_data);

    if (!status) {
        fprintf(stderr, "Error: Failed to parse xl/sharedStrings.xml\n");
        return -1;
    }

    return 0;
}

/* Styles parsing state */
typedef struct {
    xlsx2csvConverter *conv;
    bool               in_num_fmts;
    bool               in_num_fmt;
    bool               in_cell_xfs;
    bool               in_xf;
    int                format_idx;
    int                xf_idx;
    char              *current_format_code;
    char              *current_num_fmt_id;
    char              *current_num_fmt_code;
} styles_state;

static void styles_start_element(void *userData, const XML_Char *name, const XML_Char **atts)
{
    styles_state *state = (styles_state *)userData;

    if (strcmp(name, "numFmts") == 0) {
        state->in_num_fmts = true;
        /* Count formats first - we'll do this in a separate pass */
    } else if (strcmp(name, "numFmt") == 0 && state->in_num_fmts) {
        state->in_num_fmt = true;
        free(state->current_num_fmt_id);
        free(state->current_num_fmt_code);
        state->current_num_fmt_id   = NULL;
        state->current_num_fmt_code = NULL;

        for (int i = 0; atts[i]; i += 2) {
            if (strcmp(atts[i], "numFmtId") == 0) {
                state->current_num_fmt_id = str_duplicate(atts[i + 1]);
            } else if (strcmp(atts[i], "formatCode") == 0) {
                state->current_num_fmt_code = str_duplicate(atts[i + 1]);
            }
        }
    } else if (strcmp(name, "cellXfs") == 0) {
        state->in_cell_xfs = true;
    } else if (strcmp(name, "xf") == 0 && state->in_cell_xfs) {
        state->in_xf = true;
        free(state->current_format_code);
        state->current_format_code = NULL;

        for (int i = 0; atts[i]; i += 2) {
            if (strcmp(atts[i], "numFmtId") == 0) {
                state->current_format_code = str_duplicate(atts[i + 1]);
            }
        }
    }
}

static void styles_end_element(void *userData, const XML_Char *name)
{
    styles_state *state = (styles_state *)userData;

    if (strcmp(name, "numFmts") == 0) {
        state->in_num_fmts = false;
    } else if (strcmp(name, "numFmt") == 0 && state->in_num_fmt) {
        if (state->format_idx < state->conv->styles.format_count) {
            if (state->current_num_fmt_id) {
                state->conv->styles.formats[state->format_idx].id = atoi(state->current_num_fmt_id);
            }
            if (state->current_num_fmt_code) {
                state->conv->styles.formats[state->format_idx].format_code =
                    state->current_num_fmt_code;
                state->current_num_fmt_code = NULL;
            }
            state->format_idx++;
        }
        state->in_num_fmt = false;
    } else if (strcmp(name, "cellXfs") == 0) {
        state->in_cell_xfs = false;
    } else if (strcmp(name, "xf") == 0 && state->in_xf) {
        if (state->xf_idx < state->conv->styles.cell_xfs_count) {
            if (state->current_format_code) {
                state->conv->styles.cell_xfs[state->xf_idx] = atoi(state->current_format_code);
            }
            state->xf_idx++;
        }
        state->in_xf = false;
    }
}

/* Parse Styles XML */
int parse_styles(xlsx2csvConverter *conv)
{
    char *xml_data = zip_read_file_to_string(conv->zip_handle, "xl/styles.xml");
    if (!xml_data) {
        /* No styles is valid */
        conv->styles.formats        = NULL;
        conv->styles.format_count   = 0;
        conv->styles.cell_xfs       = NULL;
        conv->styles.cell_xfs_count = 0;
        return 0;
    }

    /* First pass: count formats and xfs */
    XML_Parser parser = XML_ParserCreate(NULL);
    if (!parser) {
        free(xml_data);
        return -1;
    }

    int  format_count = 0;
    int  xf_count     = 0;
    bool in_num_fmts  = false;
    bool in_cell_xfs  = false;

    count_styles_state_t count_state = {&format_count, &xf_count, &in_num_fmts, &in_cell_xfs};
    XML_SetUserData(parser, &count_state);
    XML_SetElementHandler(parser, count_styles_start, count_styles_end);
    int status = XML_Parse(parser, xml_data, (int)strlen(xml_data), 1);
    XML_ParserFree(parser);

    if (!status) {
        free(xml_data);
        fprintf(stderr, "Error: Failed to parse xl/styles.xml\n");
        return -1;
    }

    /* Allocate arrays */
    conv->styles.formats        = calloc((size_t)format_count, sizeof(numFormat));
    conv->styles.format_count   = format_count;
    conv->styles.cell_xfs       = calloc((size_t)xf_count, sizeof(int));
    conv->styles.cell_xfs_count = xf_count;

    /* Second pass: parse formats and xfs */
    parser = XML_ParserCreate(NULL);
    if (!parser) {
        free(xml_data);
        return -1;
    }

    styles_state state = {0};
    state.conv         = conv;
    XML_SetUserData(parser, &state);
    XML_SetElementHandler(parser, styles_start_element, styles_end_element);
    status = XML_Parse(parser, xml_data, (int)strlen(xml_data), 1);
    XML_ParserFree(parser);
    free(xml_data);

    if (!status) {
        fprintf(stderr, "Error: Failed to parse xl/styles.xml\n");
        return -1;
    }

    return 0;
}

/* Worksheet parsing state */
typedef struct {
    xlsx2csvConverter *conv;
    FILE              *outfile;
    csvWriter         *writer;
    int                global_max_col;
    int                last_row;
    bool               in_sheet_data;
    bool               in_row;
    bool               in_cell;
    bool               in_v;
    bool               in_is;
    bool               in_t;
    int                current_row_num;
    bool               current_row_hidden;
    char              *current_cell_ref;
    char              *current_cell_type;
    char              *current_cell_style;
    char              *current_cell_value;
    size_t             current_cell_value_len;
    size_t             current_cell_value_capacity;
    bool               in_inline_str;
    char              *current_dimension_ref;
#define MAX_COLS 1024
    char *cells[MAX_COLS];
    int   max_col;
} worksheet_state;

static void worksheet_start_element(void *userData, const XML_Char *name, const XML_Char **atts)
{
    worksheet_state *state = (worksheet_state *)userData;

    if (strcmp(name, "sheetData") == 0) {
        state->in_sheet_data = true;
    } else if (strcmp(name, "dimension") == 0) {
        for (int i = 0; atts[i]; i += 2) {
            if (strcmp(atts[i], "ref") == 0) {
                free(state->current_dimension_ref);
                state->current_dimension_ref = str_duplicate(atts[i + 1]);
                /* Parse dimension to get max column */
                char *colon = strchr(state->current_dimension_ref, ':');
                if (colon) {
                    char col_name[10] = {0};
                    int  j            = 0;
                    for (char *p = colon + 1; *p && isalpha(*p); p++) {
                        col_name[j++] = *p;
                    }
                    col_name[j]           = '\0';
                    state->global_max_col = column_name_to_index(col_name);
                }
            }
        }
    } else if (strcmp(name, "row") == 0 && state->in_sheet_data) {
        state->in_row             = true;
        state->current_row_num    = state->last_row + 1;
        state->current_row_hidden = false;
        state->max_col            = -1;
        memset(state->cells, 0, sizeof(state->cells));

        for (int i = 0; atts[i]; i += 2) {
            if (strcmp(atts[i], "r") == 0) {
                state->current_row_num = atoi(atts[i + 1]);
            } else if (strcmp(atts[i], "hidden") == 0) {
                if (strcmp(atts[i + 1], "1") == 0 || strcmp(atts[i + 1], "true") == 0) {
                    state->current_row_hidden = true;
                }
            }
        }

        /* Write empty rows if skip_empty_lines is false */
        if (!state->conv->options.skip_empty_lines) {
            for (int i = state->last_row + 1; i < state->current_row_num; i++) {
                fputs(state->conv->options.lineterminator, state->outfile);
            }
        }
        state->last_row = state->current_row_num;
    } else if (strcmp(name, "c") == 0 && state->in_row) {
        state->in_cell = true;
        free(state->current_cell_ref);
        free(state->current_cell_type);
        free(state->current_cell_style);
        free(state->current_cell_value);
        state->current_cell_ref            = NULL;
        state->current_cell_type           = NULL;
        state->current_cell_style          = NULL;
        state->current_cell_value          = NULL;
        state->current_cell_value_len      = 0;
        state->current_cell_value_capacity = 0;
        state->in_v                        = false;
        state->in_is                       = false;
        state->in_t                        = false;
        state->in_inline_str               = false;

        for (int i = 0; atts[i]; i += 2) {
            if (strcmp(atts[i], "r") == 0) {
                state->current_cell_ref = str_duplicate(atts[i + 1]);
            } else if (strcmp(atts[i], "t") == 0) {
                state->current_cell_type = str_duplicate(atts[i + 1]);
            } else if (strcmp(atts[i], "s") == 0) {
                state->current_cell_style = str_duplicate(atts[i + 1]);
            }
        }
    } else if (strcmp(name, "v") == 0 && state->in_cell) {
        state->in_v = true;
    } else if (strcmp(name, "is") == 0 && state->in_cell) {
        state->in_is         = true;
        state->in_inline_str = true;
    } else if (strcmp(name, "t") == 0) {
        /* <t> can be inside <is> (inlineStr) or inside <v> (shared string reference) */
        if (state->in_is && state->in_cell) {
            state->in_t = true;
        } else if (state->in_v && state->in_cell) {
            state->in_t = true;
        }
    }
}

static void worksheet_end_element(void *userData, const XML_Char *name)
{
    worksheet_state *state = (worksheet_state *)userData;

    if (strcmp(name, "sheetData") == 0) {
        state->in_sheet_data = false;
    } else if (strcmp(name, "row") == 0 && state->in_row) {
        /* Check if row is hidden */
        if (state->current_row_hidden && state->conv->options.skip_hidden_rows) {
            /* Free cells and skip */
            for (int i = 0; i < MAX_COLS; i++) {
                free(state->cells[i]);
                state->cells[i] = NULL;
            }
            state->in_row = false;
            return;
        }

        /* Process row */
        /* Check if row is empty */
        bool is_empty = true;
        for (int i = 0; i <= state->max_col; i++) {
            if (state->cells[i] && state->cells[i][0] != '\0') {
                is_empty = false;
                break;
            }
        }

        /* Check for date format error */
        if (state->conv->has_date_error) {
            for (int i = 0; i < MAX_COLS; i++) {
                free(state->cells[i]);
                state->cells[i] = NULL;
            }
            state->in_row = false;
            return;
        }

        /* Write row if not empty or if we're not skipping empty lines */
        if (!is_empty || !state->conv->options.skip_empty_lines) {
            int output_max_col = state->max_col;
            if (state->conv->options.skip_trailing_columns) {
                while (output_max_col >= 0 &&
                       (!state->cells[output_max_col] || state->cells[output_max_col][0] == '\0')) {
                    output_max_col--;
                }
            } else {
                if (state->global_max_col > output_max_col) {
                    output_max_col = state->global_max_col;
                }
            }

            csv_writer_reset_row(state->writer);
            csv_writer_set_field_count(state->writer, output_max_col + 1);

            if (output_max_col >= 0) {
                for (int i = 0; i <= output_max_col; i++) {
                    csv_write_field(state->writer, state->cells[i] ? state->cells[i] : "");
                }
            }
            fputs(state->conv->options.lineterminator, state->outfile);
        }

        /* Free cells */
        for (int i = 0; i < MAX_COLS; i++) {
            free(state->cells[i]);
            state->cells[i] = NULL;
        }

        state->in_row = false;
    } else if (strcmp(name, "c") == 0 && state->in_cell) {
        /* Extract column from reference */
        int col_index = 0;
        if (state->current_cell_ref) {
            char col_name[10] = {0};
            int  j            = 0;
            for (const char *p = state->current_cell_ref; *p && isalpha(*p); p++) {
                col_name[j++] = *p;
            }
            col_name[j] = '\0';
            col_index   = column_name_to_index(col_name);
        }

        /* Get cell value */
        char *value = NULL;
        if (state->current_cell_type && strcmp(state->current_cell_type, "inlineStr") == 0) {
            /* inlineStr: value should be collected from <is><t>...</t></is> */
            if (state->current_cell_value) {
                value = str_duplicate(state->current_cell_value);
            } else {
                value = str_duplicate("");
            }
        } else if (state->current_cell_value) {
            value = format_cell_value(state->current_cell_value,
                                      state->current_cell_type,
                                      state->current_cell_style,
                                      state->conv);
        }

        /* Store value in correct column */
        if (col_index >= 0 && col_index < MAX_COLS) {
            if (value) {
                state->cells[col_index] = value;
            } else {
                state->cells[col_index] = str_duplicate("");
            }
            if (col_index > state->max_col) {
                state->max_col = col_index;
            }
        } else if (value) {
            free(value);
        }

        free(state->current_cell_ref);
        free(state->current_cell_type);
        free(state->current_cell_style);
        free(state->current_cell_value);
        state->current_cell_ref            = NULL;
        state->current_cell_type           = NULL;
        state->current_cell_style          = NULL;
        state->current_cell_value          = NULL;
        state->current_cell_value_len      = 0;
        state->current_cell_value_capacity = 0;

        state->in_cell = false;
    } else if (strcmp(name, "v") == 0) {
        state->in_v = false;
    } else if (strcmp(name, "is") == 0) {
        state->in_is         = false;
        state->in_inline_str = false;
    } else if (strcmp(name, "t") == 0) {
        state->in_t = false;
    }
}

static void worksheet_char_data(void *userData, const XML_Char *s, int len)
{
    worksheet_state *state = (worksheet_state *)userData;

    /* Collect text data:
     * - If in <v> node (direct text content, no <t> wrapper)
     * - If in <is><t> (inline string with <t> wrapper)
     */
    if ((state->in_v && state->in_cell) || (state->in_t && state->in_is && state->in_cell)) {
        size_t new_len = state->current_cell_value_len + (size_t)len;
        if (new_len >= state->current_cell_value_capacity) {
            state->current_cell_value_capacity = new_len + 256;
            state->current_cell_value =
                realloc(state->current_cell_value, state->current_cell_value_capacity + 1);
            if (!state->current_cell_value) {
                return;
            }
        }
        memcpy(state->current_cell_value + state->current_cell_value_len, s, (size_t)len);
        state->current_cell_value_len      = new_len;
        state->current_cell_value[new_len] = '\0';
    }
}

/* Parse worksheet and convert to CSV */
int parse_worksheet(xlsx2csvConverter *conv, int sheet_index, FILE *outfile)
{
    if (!conv || !outfile) {
        return -1;
    }

    /* Build worksheet filename */
    char filename[256];
    snprintf(filename, sizeof(filename), "xl/worksheets/sheet%d.xml", sheet_index);

    char *xml_data = zip_read_file_to_string(conv->zip_handle, filename);
    if (!xml_data) {
        fprintf(stderr, "Error: Could not read %s\n", filename);
        return -1;
    }

    XML_Parser parser = XML_ParserCreate(NULL);
    if (!parser) {
        free(xml_data);
        return -1;
    }

    worksheet_state state = {0};
    state.conv            = conv;
    state.outfile         = outfile;
    state.writer          = csv_writer_create(outfile, &conv->options);
    state.last_row        = 0;
    state.global_max_col  = -1;

    if (!state.writer) {
        XML_ParserFree(parser);
        free(xml_data);
        return -1;
    }

    XML_SetUserData(parser, &state);
    XML_SetElementHandler(parser, worksheet_start_element, worksheet_end_element);
    XML_SetCharacterDataHandler(parser, worksheet_char_data);

    int status = XML_Parse(parser, xml_data, (int)strlen(xml_data), 1);
    XML_ParserFree(parser);
    free(xml_data);

    csv_writer_free(state.writer);

    free(state.current_dimension_ref);
    free(state.current_cell_ref);
    free(state.current_cell_type);
    free(state.current_cell_style);
    free(state.current_cell_value);

    if (!status) {
        fprintf(stderr, "Error: Failed to parse %s\n", filename);
        return -1;
    }

    return 0;
}
