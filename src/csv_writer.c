#include <csv.h>
#include <stdlib.h>
#include <string.h>
#include "xlsx2csv.h"

/* CSV Writer structure */
struct csvWriter {
    FILE        *fp;
    xlsxOptions *options;
    int          field_index;
    int          field_count; /* Total fields in current row */
};

/* Create CSV writer */
csvWriter *csv_writer_create(FILE *fp, xlsxOptions *options)
{
    if (!fp || !options) {
        return NULL;
    }

    csvWriter *writer = calloc(1, sizeof(csvWriter));
    if (!writer) {
        return NULL;
    }

    writer->fp          = fp;
    writer->options     = options;
    writer->field_index = 0;
    writer->field_count = 0;

    return writer;
}

/* Free CSV writer */
void csv_writer_free(csvWriter *writer)
{
    if (writer) {
        free(writer);
    }
}

/* Check if field needs quoting based on quoting mode */
static bool needs_quoting(const char *field, xlsxOptions *options, int field_count, int field_index)
{
    if (!field) {
        return false;
    }

    switch (options->quoting) {
        case QUOTE_ALL:
            return true;

        case QUOTE_NONNUMERIC:
            /* Python's csv.QUOTE_NONNUMERIC quotes all fields that are strings
             * Since xlsx2csv outputs everything as strings, this means ALL fields
             * are quoted in QUOTE_NONNUMERIC mode (matching Python behavior) */
            return true;

        case QUOTE_MINIMAL:
            /* Empty strings: only quote if it's the ONLY field in the row */
            if (field[0] == '\0') {
                return (field_count == 1 && field_index == 0);
            }
            /* Quote if field contains delimiter, quote, or line breaks */
            if (strchr(field, options->delimiter) != NULL) {
                return true;
            }
            if (strchr(field, '"') != NULL) {
                return true;
            }
            if (strchr(field, '\n') != NULL || strchr(field, '\r') != NULL) {
                return true;
            }
            return false;

        case QUOTE_NONE:
            return false;

        default:
            return false;
    }
}

/* Write a single field */
int csv_write_field(csvWriter *writer, const char *field)
{
    if (!writer || !writer->fp) {
        return -1;
    }

    /* Write delimiter if not first field */
    if (writer->field_index > 0) {
        fputc(writer->options->delimiter, writer->fp);
    }

    /* Handle NULL field */
    if (!field) {
        field = "";
    }

    /* Check if empty field needs quoting */
    if (field[0] == '\0') {
        if (needs_quoting(field, writer->options, writer->field_count, writer->field_index)) {
            fputs("\"\"", writer->fp);
        }
        writer->field_index++;
        return 0;
    }

    /* Apply line break handling */
    char *processed_field = NULL;
    if (writer->options->no_line_breaks) {
        /* Replace line breaks with spaces */
        processed_field = str_duplicate(field);
        if (processed_field) {
            for (char *p = processed_field; *p; p++) {
                if (*p == '\r' || *p == '\n' || *p == '\t') {
                    *p = ' ';
                }
            }
        }
        field = processed_field;
    } else if (writer->options->escape_strings) {
        /* Escape control characters */
        size_t len         = strlen(field);
        size_t escaped_len = len * 2 + 1; /* Worst case */
        processed_field    = malloc(escaped_len);
        if (processed_field) {
            char *dst = processed_field;
            for (const char *src = field; *src; src++) {
                if (*src == '\r') {
                    *dst++ = '\\';
                    *dst++ = 'r';
                } else if (*src == '\n') {
                    *dst++ = '\\';
                    *dst++ = 'n';
                } else if (*src == '\t') {
                    *dst++ = '\\';
                    *dst++ = 't';
                } else {
                    *dst++ = *src;
                }
            }
            *dst  = '\0';
            field = processed_field;
        }
    }

    bool quote = needs_quoting(field, writer->options, writer->field_count, writer->field_index);

    if (quote) {
        fputc('"', writer->fp);
    }

    /* Write field content, escaping quotes if needed */
    if (writer->options->quoting != QUOTE_NONE) {
        /* Use libcsv to write the field content (handles quote escaping) */
        for (const char *p = field; *p; p++) {
            if (*p == '"') {
                fputc('"', writer->fp); /* Double the quote */
                fputc('"', writer->fp);
            } else {
                fputc(*p, writer->fp);
            }
        }
    } else {
        /* No quoting mode - just write as-is */
        fputs(field, writer->fp);
    }

    if (quote) {
        fputc('"', writer->fp);
    }

    if (processed_field) {
        free(processed_field);
    }

    writer->field_index++;
    return 0;
}

/* Write a complete row */
int csv_write_row(csvWriter *writer, char **fields, int field_count)
{
    if (!writer || !writer->fp) {
        return -1;
    }

    /* Reset field index and set field count */
    writer->field_index = 0;
    writer->field_count = field_count;

    /* Write all fields */
    for (int i = 0; i < field_count; i++) {
        if (csv_write_field(writer, fields[i]) < 0) {
            return -1;
        }
    }

    /* Write line terminator */
    fputs(writer->options->lineterminator, writer->fp);

    return 0;
}

/* Reset row (for manual field writing) */
void csv_writer_reset_row(csvWriter *writer)
{
    if (writer) {
        writer->field_index = 0;
    }
}

/* Set field count for current row */
void csv_writer_set_field_count(csvWriter *writer, int count)
{
    if (writer) {
        writer->field_count = count;
    }
}
