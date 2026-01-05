/* Standard library headers */
#include <ctype.h>
#include <stdlib.h>
#include <string.h>

/* Third-party library headers */
#include <libxml/parser.h>
#include <libxml/tree.h>
#include <libxml/xpath.h>

/* Project headers */
#include "csv_writer.h"
#include "format_handler.h"
#include "utils.h"
#include "xlsx2csv.h"
#include "xml_parser.h"
#include "zip_reader.h"

/* Parse Content Types XML */
int parse_content_types(xlsx2csvConverter *conv)
{
    char *xml_data = zip_read_file_to_string(conv->zip_handle, "[Content_Types].xml");
    if (!xml_data) {
        fprintf(stderr, "Error: Could not read [Content_Types].xml\n");
        return -1;
    }

    xmlDocPtr doc = xmlReadMemory(xml_data, strlen(xml_data), NULL, NULL, 0);
    free(xml_data);

    if (!doc) {
        fprintf(stderr, "Error: Failed to parse [Content_Types].xml\n");
        return -1;
    }

    /* Content types parsing not critical for basic functionality */
    xmlFreeDoc(doc);
    return 0;
}

/* Parse Workbook XML */
int parse_workbook(xlsx2csvConverter *conv)
{
    char *xml_data = zip_read_file_to_string(conv->zip_handle, "xl/workbook.xml");
    if (!xml_data) {
        fprintf(stderr, "Error: Could not read xl/workbook.xml\n");
        return -1;
    }

    xmlDocPtr doc = xmlReadMemory(xml_data, strlen(xml_data), NULL, NULL, 0);
    free(xml_data);

    if (!doc) {
        fprintf(stderr, "Error: Failed to parse xl/workbook.xml\n");
        return -1;
    }

    xmlNodePtr root = xmlDocGetRootElement(doc);
    if (!root) {
        xmlFreeDoc(doc);
        return -1;
    }

    /* Check for date1904 */
    conv->workbook.date1904 = false;
    for (xmlNodePtr node = root->children; node; node = node->next) {
        if (node->type == XML_ELEMENT_NODE &&
            xmlStrcmp(node->name, (const xmlChar *)"workbookPr") == 0) {
            xmlChar *date1904 = xmlGetProp(node, (const xmlChar *)"date1904");
            if (date1904) {
                if (xmlStrcmp(date1904, (const xmlChar *)"true") == 0 ||
                    xmlStrcmp(date1904, (const xmlChar *)"1") == 0) {
                    conv->workbook.date1904 = true;
                }
                xmlFree(date1904);
            }
        }
    }

    /* Find sheets */
    xmlNodePtr sheets_node = NULL;
    for (xmlNodePtr node = root->children; node; node = node->next) {
        if (node->type == XML_ELEMENT_NODE &&
            xmlStrcmp(node->name, (const xmlChar *)"sheets") == 0) {
            sheets_node = node;
            break;
        }
    }

    if (!sheets_node) {
        xmlFreeDoc(doc);
        return -1;
    }

    /* Count sheets */
    int sheet_count = 0;
    for (xmlNodePtr node = sheets_node->children; node; node = node->next) {
        if (node->type == XML_ELEMENT_NODE &&
            xmlStrcmp(node->name, (const xmlChar *)"sheet") == 0) {
            sheet_count++;
        }
    }

    /* Allocate sheet array */
    conv->workbook.sheets      = calloc((size_t)sheet_count, sizeof(sheetInfo));
    conv->workbook.sheet_count = sheet_count;

    /* Parse sheets */
    int idx = 0;
    for (xmlNodePtr node = sheets_node->children; node; node = node->next) {
        if (node->type == XML_ELEMENT_NODE &&
            xmlStrcmp(node->name, (const xmlChar *)"sheet") == 0) {

            xmlChar *name  = xmlGetProp(node, (const xmlChar *)"name");
            xmlChar *rid   = xmlGetProp(node, (const xmlChar *)"r:id");
            xmlChar *state = xmlGetProp(node, (const xmlChar *)"state");

            if (name) {
                conv->workbook.sheets[idx].name = str_duplicate((char *)name);
                xmlFree(name);
            }
            if (rid) {
                conv->workbook.sheets[idx].relation_id = str_duplicate((char *)rid);
                xmlFree(rid);
            }
            if (state) {
                conv->workbook.sheets[idx].state = str_duplicate((char *)state);
                xmlFree(state);
            }

            conv->workbook.sheets[idx].index = idx + 1;
            idx++;
        }
    }

    xmlFreeDoc(doc);
    return 0;
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

    xmlDocPtr doc = xmlReadMemory(xml_data, strlen(xml_data), NULL, NULL, 0);
    free(xml_data);

    if (!doc) {
        fprintf(stderr, "Error: Failed to parse xl/sharedStrings.xml\n");
        return -1;
    }

    xmlNodePtr root = xmlDocGetRootElement(doc);
    if (!root) {
        xmlFreeDoc(doc);
        return -1;
    }

    /* Count strings */
    int count = 0;
    for (xmlNodePtr si = root->children; si; si = si->next) {
        if (si->type == XML_ELEMENT_NODE && xmlStrcmp(si->name, (const xmlChar *)"si") == 0) {
            count++;
        }
    }

    /* Allocate string array */
    conv->shared_strings.strings = calloc((size_t)count, sizeof(char *));
    conv->shared_strings.count   = count;

    /* Parse strings */
    int idx = 0;
    for (xmlNodePtr si = root->children; si; si = si->next) {
        if (si->type == XML_ELEMENT_NODE && xmlStrcmp(si->name, (const xmlChar *)"si") == 0) {

            /* Find <t> element */
            xmlNodePtr t_node = NULL;
            for (xmlNodePtr child = si->children; child; child = child->next) {
                if (child->type == XML_ELEMENT_NODE &&
                    xmlStrcmp(child->name, (const xmlChar *)"t") == 0) {
                    t_node = child;
                    break;
                }
            }

            if (t_node) {
                xmlChar *content = xmlNodeGetContent(t_node);
                if (content) {
                    conv->shared_strings.strings[idx] = str_duplicate((char *)content);
                    xmlFree(content);
                } else {
                    conv->shared_strings.strings[idx] = str_duplicate("");
                }
            } else {
                conv->shared_strings.strings[idx] = str_duplicate("");
            }

            idx++;
        }
    }

    xmlFreeDoc(doc);
    return 0;
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

    xmlDocPtr doc = xmlReadMemory(xml_data, strlen(xml_data), NULL, NULL, 0);
    free(xml_data);

    if (!doc) {
        fprintf(stderr, "Error: Failed to parse xl/styles.xml\n");
        return -1;
    }

    xmlNodePtr root = xmlDocGetRootElement(doc);
    if (!root) {
        xmlFreeDoc(doc);
        return -1;
    }

    /* Parse numFmts */
    for (xmlNodePtr node = root->children; node; node = node->next) {
        if (node->type == XML_ELEMENT_NODE &&
            xmlStrcmp(node->name, (const xmlChar *)"numFmts") == 0) {

            /* Count formats */
            int count = 0;
            for (xmlNodePtr fmt = node->children; fmt; fmt = fmt->next) {
                if (fmt->type == XML_ELEMENT_NODE) {
                    count++;
                }
            }

            conv->styles.formats      = calloc((size_t)count, sizeof(numFormat));
            conv->styles.format_count = count;

            /* Parse formats */
            int idx = 0;
            for (xmlNodePtr fmt = node->children; fmt; fmt = fmt->next) {
                if (fmt->type == XML_ELEMENT_NODE) {
                    xmlChar *id   = xmlGetProp(fmt, (const xmlChar *)"numFmtId");
                    xmlChar *code = xmlGetProp(fmt, (const xmlChar *)"formatCode");

                    if (id) {
                        conv->styles.formats[idx].id = atoi((char *)id);
                        xmlFree(id);
                    }
                    if (code) {
                        conv->styles.formats[idx].format_code = str_duplicate((char *)code);
                        xmlFree(code);
                    }
                    idx++;
                }
            }
        }
    }

    /* Parse cellXfs */
    for (xmlNodePtr node = root->children; node; node = node->next) {
        if (node->type == XML_ELEMENT_NODE &&
            xmlStrcmp(node->name, (const xmlChar *)"cellXfs") == 0) {

            /* Count xfs */
            int count = 0;
            for (xmlNodePtr xf = node->children; xf; xf = xf->next) {
                if (xf->type == XML_ELEMENT_NODE) {
                    count++;
                }
            }

            conv->styles.cell_xfs       = calloc((size_t)count, sizeof(int));
            conv->styles.cell_xfs_count = count;

            /* Parse xfs */
            int idx = 0;
            for (xmlNodePtr xf = node->children; xf; xf = xf->next) {
                if (xf->type == XML_ELEMENT_NODE) {
                    xmlChar *fmt_id = xmlGetProp(xf, (const xmlChar *)"numFmtId");
                    if (fmt_id) {
                        conv->styles.cell_xfs[idx] = atoi((char *)fmt_id);
                        xmlFree(fmt_id);
                    }
                    idx++;
                }
            }
        }
    }

    xmlFreeDoc(doc);
    return 0;
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

    xmlDocPtr doc = xmlReadMemory(xml_data, strlen(xml_data), NULL, NULL, 0);
    free(xml_data);

    if (!doc) {
        fprintf(stderr, "Error: Failed to parse %s\n", filename);
        return -1;
    }

    xmlNodePtr root = xmlDocGetRootElement(doc);
    if (!root) {
        xmlFreeDoc(doc);
        return -1;
    }

    /* Find sheetData node and parse dimension */
    xmlNodePtr sheetData      = NULL;
    int        global_max_col = -1;

    for (xmlNodePtr node = root->children; node; node = node->next) {
        if (node->type == XML_ELEMENT_NODE) {
            if (xmlStrcmp(node->name, (const xmlChar *)"sheetData") == 0) {
                sheetData = node;
            } else if (xmlStrcmp(node->name, (const xmlChar *)"dimension") == 0) {
                /* Parse dimension to get max column */
                xmlChar *ref = xmlGetProp(node, (const xmlChar *)"ref");
                if (ref) {
                    /* ref format: "A1:D6" or just "A1" */
                    char *colon = strchr((char *)ref, ':');
                    if (colon) {
                        /* Extract column from end reference (e.g., "D6" -> "D") */
                        char col_name[10] = {0};
                        int  i            = 0;
                        for (char *p = colon + 1; *p && isalpha(*p); p++) {
                            col_name[i++] = *p;
                        }
                        col_name[i]    = '\0';
                        global_max_col = column_name_to_index(col_name);
                    }
                    xmlFree(ref);
                }
            }
        }
    }

    if (!sheetData) {
        xmlFreeDoc(doc);
        return 0; /* Empty sheet */
    }

    /* Create CSV writer */
    csvWriter *writer = csv_writer_create(outfile, &conv->options);
    if (!writer) {
        xmlFreeDoc(doc);
        return -1;
    }

    /* Process rows */
    int last_row = 0;
    for (xmlNodePtr row = sheetData->children; row; row = row->next) {
        if (row->type != XML_ELEMENT_NODE || xmlStrcmp(row->name, (const xmlChar *)"row") != 0) {
            continue;
        }

        /* Check if row is hidden */
        xmlChar *hidden = xmlGetProp(row, (const xmlChar *)"hidden");
        if (hidden && conv->options.skip_hidden_rows) {
            if (xmlStrcmp(hidden, (const xmlChar *)"1") == 0 ||
                xmlStrcmp(hidden, (const xmlChar *)"true") == 0) {
                xmlFree(hidden);
                continue;
            }
        }
        if (hidden) {
            xmlFree(hidden);
        }

        /* Get row number */
        xmlChar *row_num_str = xmlGetProp(row, (const xmlChar *)"r");
        int      row_num     = row_num_str ? atoi((char *)row_num_str) : last_row + 1;
        if (row_num_str) {
            xmlFree(row_num_str);
        }

        /* Write empty rows if skip_empty_lines is false */
        if (!conv->options.skip_empty_lines) {
            for (int i = last_row + 1; i < row_num; i++) {
                fputs(conv->options.lineterminator, outfile);
            }
        }
        last_row = row_num;

/* Collect cells */
#define MAX_COLS 1024
        char *cells[MAX_COLS] = {0};
        int   max_col         = -1;

        for (xmlNodePtr cell = row->children; cell; cell = cell->next) {
            if (cell->type != XML_ELEMENT_NODE) {
                continue;
            }
            if (xmlStrcmp(cell->name, (const xmlChar *)"c") != 0) {
                continue;
            }

            /* Get cell reference (e.g., "A1", "B2") */
            xmlChar *ref        = xmlGetProp(cell, (const xmlChar *)"r");
            xmlChar *type_attr  = xmlGetProp(cell, (const xmlChar *)"t");
            xmlChar *style_attr = xmlGetProp(cell, (const xmlChar *)"s");

            /* Extract column from reference */
            int col_index = 0;
            if (ref) {
                char col_name[10] = {0};
                int  i            = 0;
                for (const xmlChar *p = ref; *p && isalpha(*p); p++) {
                    col_name[i++] = *p;
                }
                col_name[i] = '\0';
                col_index   = column_name_to_index(col_name);
            }

            /* Find value */
            xmlNodePtr v_node  = NULL;
            xmlNodePtr is_node = NULL;
            for (xmlNodePtr child = cell->children; child; child = child->next) {
                if (child->type == XML_ELEMENT_NODE) {
                    if (xmlStrcmp(child->name, (const xmlChar *)"v") == 0) {
                        v_node = child;
                    } else if (xmlStrcmp(child->name, (const xmlChar *)"is") == 0) {
                        is_node = child;
                    }
                }
            }

            /* Get cell value */
            char *value = NULL;
            if (type_attr && xmlStrcmp(type_attr, (const xmlChar *)"inlineStr") == 0 && is_node) {
                /* Handle inline string: look for <t> inside <is> */
                bool found = false;
                for (xmlNodePtr t = is_node->children; t; t = t->next) {
                    if (t->type == XML_ELEMENT_NODE &&
                        xmlStrcmp(t->name, (const xmlChar *)"t") == 0) {
                        xmlChar *content = xmlNodeGetContent(t);
                        if (content) {
                            value = str_duplicate((char *)content);
                            xmlFree(content);
                        }
                        found = true;
                        break;
                    }
                }
                /* If inlineStr but no <t> found, it's an empty string */
                if (!found) {
                    value = str_duplicate("");
                }
            } else if (type_attr && xmlStrcmp(type_attr, (const xmlChar *)"inlineStr") == 0) {
                /* inlineStr without <is> node - empty string */
                value = str_duplicate("");
            } else if (v_node) {
                xmlChar *content = xmlNodeGetContent(v_node);
                if (content) {
                    value = format_cell_value(
                        (char *)content, (char *)type_attr, (char *)style_attr, conv);
                    xmlFree(content);
                }
            }

            /* Store value in correct column */
            if (col_index >= 0 && col_index < MAX_COLS) {
                if (value) {
                    cells[col_index] = value;
                } else {
                    cells[col_index] = str_duplicate("");
                }
                if (col_index > max_col) {
                    max_col = col_index;
                }
            } else if (value) {
                free(value); /* Free if column out of range */
            }

            if (ref)
                xmlFree(ref);
            if (type_attr)
                xmlFree(type_attr);
            if (style_attr)
                xmlFree(style_attr);
        }

        /* Check if row is empty */
        bool is_empty = true;
        for (int i = 0; i <= max_col; i++) {
            if (cells[i] && cells[i][0] != '\0') {
                is_empty = false;
                break;
            }
        }

        /* Check for date format error - stop processing before writing this row */
        if (conv->has_date_error) {
            /* Free cells and break */
            for (int i = 0; i < MAX_COLS; i++) {
                free(cells[i]);
            }
            break;
        }

        /* Write row if not empty or if we're not skipping empty lines */
        if (!is_empty || !conv->options.skip_empty_lines) {
            /* Adjust max_col if skip_trailing_columns is enabled */
            int output_max_col = max_col;
            if (conv->options.skip_trailing_columns) {
                while (output_max_col >= 0 &&
                       (!cells[output_max_col] || cells[output_max_col][0] == '\0')) {
                    output_max_col--;
                }
            } else {
                /* Use global max column if available */
                if (global_max_col > output_max_col) {
                    output_max_col = global_max_col;
                }
            }

            /* Reset writer field index for new row */
            csv_writer_reset_row(writer);

            /* Set field count for this row (needed for proper empty field quoting) */
            csv_writer_set_field_count(writer, output_max_col + 1);

            /* Write cells - ensure we only write up to output_max_col + 1 cells */
            if (output_max_col >= 0) {
                for (int i = 0; i <= output_max_col; i++) {
                    csv_write_field(writer, cells[i] ? cells[i] : "");
                }
            } else {
                /* No cells in row, but still need to output empty line if not skipping */
                if (!conv->options.skip_empty_lines) {
                    /* Output nothing for empty row */
                }
            }
            fputs(conv->options.lineterminator, outfile);
        }

        /* Free cells */
        for (int i = 0; i < MAX_COLS; i++) {
            free(cells[i]);
        }
    }

    csv_writer_free(writer);
    xmlFreeDoc(doc);

    return 0;
}
