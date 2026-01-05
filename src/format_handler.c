#include <ctype.h>
#include <math.h>
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include "xlsx2csv.h"

/* Format mappings (from Python version) */
static const struct {
    const char *format;
    formatType  type;
} format_map[] = {
    {"general",               FORMAT_FLOAT     },
    {"0",                     FORMAT_FLOAT     },
    {"0.00",                  FORMAT_FLOAT     },
    {"#,##0",                 FORMAT_FLOAT     },
    {"#,##0.00",              FORMAT_FLOAT     },
    {"0%",                    FORMAT_PERCENTAGE},
    {"0.00%",                 FORMAT_PERCENTAGE},
    {"0.00e+00",              FORMAT_FLOAT     },
    {"mm-dd-yy",              FORMAT_DATE      },
    {"d-mmm-yy",              FORMAT_DATE      },
    {"d-mmm",                 FORMAT_DATE      },
    {"mmm-yy",                FORMAT_DATE      },
    {"h:mm am/pm",            FORMAT_DATE      },
    {"h:mm:ss am/pm",         FORMAT_DATE      },
    {"h:mm",                  FORMAT_TIME      },
    {"h:mm:ss",               FORMAT_TIME      },
    {"m/d/yy h:mm",           FORMAT_DATE      },
    {"mm:ss",                 FORMAT_TIME      },
    {"[h]:mm:ss",             FORMAT_TIME      },
    {"mmss.0",                FORMAT_TIME      },
    {"##0.0e+0",              FORMAT_FLOAT     },
    {"@",                     FORMAT_FLOAT     },
    {"yyyy\\-mm\\-dd",        FORMAT_DATE      },
    {"dd/mm/yy",              FORMAT_DATE      },
    {"hh:mm:ss",              FORMAT_TIME      },
    {"dd/mm/yy\\ hh:mm",      FORMAT_DATE      },
    {"dd/mm/yyyy hh:mm:ss",   FORMAT_DATE      },
    {"yy-mm-dd",              FORMAT_DATE      },
    {"d-mmm-yyyy",            FORMAT_DATE      },
    {"m/d/yy",                FORMAT_DATE      },
    {"m/d/yyyy",              FORMAT_DATE      },
    {"dd-mmm-yyyy",           FORMAT_DATE      },
    {"dd/mm/yyyy",            FORMAT_DATE      },
    {"mm/dd/yy h:mm am/pm",   FORMAT_DATE      },
    {"mm/dd/yy hh:mm",        FORMAT_DATE      },
    {"mm/dd/yyyy h:mm am/pm", FORMAT_DATE      },
    {"mm/dd/yyyy hh:mm:ss",   FORMAT_DATE      },
    {"yyyy-mm-dd hh:mm:ss",   FORMAT_DATE      },
    {NULL,                    FORMAT_STRING    }
};

/* Get format type from format string */
static formatType get_format_type_from_string(const char *format_str)
{
    if (!format_str) {
        return FORMAT_STRING;
    }

    for (int i = 0; format_map[i].format != NULL; i++) {
        if (strcmp(format_str, format_map[i].format) == 0) {
            return format_map[i].type;
        }
    }

    /* Check for date pattern */
    if (strstr(format_str, "y") || strstr(format_str, "m") || strstr(format_str, "d") ||
        strstr(format_str, "h") || strstr(format_str, "s")) {
        return FORMAT_DATE;
    }

    return FORMAT_FLOAT;
}

/* Get format type by style ID */
formatType get_format_type(int style_id, styleInfo *styles)
{
    if (!styles || style_id < 0 || style_id >= styles->cell_xfs_count) {
        return FORMAT_STRING;
    }

    int fmt_id = styles->cell_xfs[style_id];

    /* Check custom formats */
    for (int i = 0; i < styles->format_count; i++) {
        if (styles->formats[i].id == fmt_id) {
            return get_format_type_from_string(styles->formats[i].format_code);
        }
    }

    /* Check standard formats */
    const char *std_format = NULL;
    static const struct {
        int         id;
        const char *format;
    } standard_formats[] = {
        {0,  "general"                 },
        {1,  "0"                       },
        {2,  "0.00"                    },
        {3,  "#,##0"                   },
        {4,  "#,##0.00"                },
        {9,  "0%"                      },
        {10, "0.00%"                   },
        {11, "0.00e+00"                },
        {14, "mm-dd-yy"                },
        {15, "d-mmm-yy"                },
        {16, "d-mmm"                   },
        {17, "mmm-yy"                  },
        {18, "h:mm am/pm"              },
        {19, "h:mm:ss am/pm"           },
        {20, "h:mm"                    },
        {21, "h:mm:ss"                 },
        {22, "m/d/yy h:mm"             },
        {37, "#,##0 ;(#,##0)"          },
        {38, "#,##0 ;[red](#,##0)"     },
        {39, "#,##0.00;(#,##0.00)"     },
        {40, "#,##0.00;[red](#,##0.00)"},
        {45, "mm:ss"                   },
        {46, "[h]:mm:ss"               },
        {47, "mmss.0"                  },
        {48, "##0.0e+0"                },
        {49, "@"                       }
    };

    for (size_t i = 0; i < sizeof(standard_formats) / sizeof(standard_formats[0]); i++) {
        if (standard_formats[i].id == fmt_id) {
            std_format = standard_formats[i].format;
            break;
        }
    }

    if (std_format) {
        return get_format_type_from_string(std_format);
    }

    return FORMAT_FLOAT;
}

/* Format date value */
char *format_date(double value, const char *format, bool date1904)
{
    /* Excel date: days since 1900-01-01 (or 1904-01-01 if date1904)
     * Note: Excel has a bug where it thinks 1900 was a leap year
     * For dates after 1900-02-28, we need to subtract 1 day to compensate
     */

    /* Convert Excel serial date to year, month, day */
    int days;
    if (date1904) {
        days = (int)value;
        /* 1904-01-01 is day 0 */
        days += 24107; /* Days from 1900-01-01 to 1904-01-01 */
    } else {
        days = (int)value;
        /* Excel counts from 1900-01-01 as day 1, but has a leap year bug */
        /* For dates > 60 (after 1900-02-28), subtract 1 to compensate */
        if (days > 60) {
            days -= 1;
        }
    }

    /* Calculate date from days since 1900-01-01 */
    /* 1900-01-01 corresponds to Unix timestamp calculation base */
    int year  = 1900;
    int month = 1;
    int day   = days;

    /* Days in each month */
    int days_in_month[] = {31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31};

    while (day > 365) {
        int days_in_year = 365;
        if ((year % 4 == 0 && year % 100 != 0) || (year % 400 == 0)) {
            days_in_year = 366;
        }
        if (day > days_in_year) {
            day -= days_in_year;
            year++;
        } else {
            break;
        }
    }

    /* Check for leap year */
    bool is_leap = (year % 4 == 0 && year % 100 != 0) || (year % 400 == 0);
    if (is_leap) {
        days_in_month[1] = 29;
    }

    /* Find month and day */
    for (int m = 0; m < 12; m++) {
        if (day <= days_in_month[m]) {
            month = m + 1;
            break;
        }
        day -= days_in_month[m];
    }

    char buffer[256];
    if (format) {
        /* Use custom format - simple implementation */
        snprintf(buffer, sizeof(buffer), "%04d-%02d-%02d", year, month, day);
    } else {
        /* Default format: YYYY-MM-DD */
        snprintf(buffer, sizeof(buffer), "%04d-%02d-%02d", year, month, day);
    }

    return str_duplicate(buffer);
}

/* Format time value */
char *format_time(double value, const char *format)
{
    /* Time is fraction of day */
    int total_seconds = (int)(value * 86400);
    int hours         = total_seconds / 3600;
    int minutes       = (total_seconds % 3600) / 60;
    int seconds       = total_seconds % 60;

    char buffer[256];
    if (format) {
        snprintf(buffer, sizeof(buffer), "%02d:%02d:%02d", hours, minutes, seconds);
    } else {
        snprintf(buffer, sizeof(buffer), "%02d:%02d", hours, minutes);
    }

    return str_duplicate(buffer);
}

/* Format float value */
char *format_float(double value, const char *format, bool scifloat)
{
    char buffer[256];

    if (format) {
        snprintf(buffer, sizeof(buffer), format, value);

        /* Check if result is zero (positive or negative) */
        double result_value = atof(buffer);
        bool   is_zero      = fabs(result_value) < 1e-300;

        if (!is_zero) {
            /* Strip trailing zeros after decimal point (Python xlsx2csv behavior) */
            char *p = strchr(buffer, '.');
            if (p) {
                char *end = buffer + strlen(buffer) - 1;
                while (end > p && *end == '0') {
                    *end = '\0';
                    end--;
                }
                /* Remove trailing decimal point if no digits after it */
                if (*end == '.') {
                    *end = '\0';
                }
            }
        }

        /* Preserve negative zero: if result is "-0.00" or similar, keep as "-0" */
        if (is_zero && (value < 0.0 || signbit(value))) {
            /* Find where digits start after decimal point */
            char *p = strchr(buffer, '.');
            if (p) {
                /* Keep format like "-0.00" but change to "-0" */
                strcpy(buffer, "-0");
            } else {
                strcpy(buffer, "-0");
            }
        }
    } else if (scifloat) {
        /* Python xlsx2csv's --sci-float behavior:
         * Use regular decimal format (not scientific notation)
         * This matches Python's behavior where scifloat doesn't force scientific notation
         */
        snprintf(buffer, sizeof(buffer), "%f", value);

        /* Check if result is zero */
        double result_value = atof(buffer);
        bool   is_zero      = fabs(result_value) < 1e-300;

        if (!is_zero) {
            /* Strip trailing zeros */
            char *p = strchr(buffer, '.');
            if (p) {
                char *end = buffer + strlen(buffer) - 1;
                while (end > p && *end == '0') {
                    *end = '\0';
                    end--;
                }
                if (*end == '.') {
                    *end = '\0';
                }
            }
        }
    } else if (fabs(value - round(value)) < 1e-9) {
        /* Integer value */
        /* Check for negative zero: value is close to zero but signbit is set */
        if (fabs(value) < 1e-10) {
            /* Check if it's negative zero by checking signbit */
            if (value < 0.0 || (fabs(value) < 1e-300 && signbit(value))) {
                /* Negative zero: use %f format to preserve sign, then strip trailing zeros */
                snprintf(buffer, sizeof(buffer), "%f", value);
                /* Strip trailing zeros */
                char *p = strchr(buffer, '.');
                if (p) {
                    char *end = buffer + strlen(buffer) - 1;
                    while (end > p && *end == '0') {
                        *end = '\0';
                        end--;
                    }
                    if (*end == '.') {
                        *end = '\0';
                    }
                }
                /* If result is "-0", keep it; if "0", check original sign */
                if (strcmp(buffer, "0") == 0 && (value < 0.0 || signbit(value))) {
                    snprintf(buffer, sizeof(buffer), "-0");
                }
            } else {
                snprintf(buffer, sizeof(buffer), "0");
            }
        } else {
            snprintf(buffer, sizeof(buffer), "%.0f", value);
        }
    } else {
        /* Use %f format (default 6 decimal places), then aggressively strip trailing zeros
         * This matches Python's behavior: ("%f" % data).rstrip('0').rstrip('.')
         */
        snprintf(buffer, sizeof(buffer), "%f", value);

        /* Strip trailing zeros after decimal point */
        char *p = strchr(buffer, '.');
        if (p) {
            char *end = buffer + strlen(buffer) - 1;
            while (end > p && *end == '0') {
                *end = '\0';
                end--;
            }
            /* Remove trailing decimal point if no digits after it */
            if (*end == '.') {
                *end = '\0';
            }
        }
    }

    return str_duplicate(buffer);
}

/* Main cell formatting function */
char *format_cell_value(const char        *value,
                        const char        *type_attr,
                        const char        *style_attr,
                        xlsx2csvConverter *conv)
{
    if (!value) {
        return str_duplicate("");
    }

    /* Handle shared string */
    if (type_attr && strcmp(type_attr, "s") == 0) {
        int index = atoi(value);
        if (index >= 0 && index < conv->shared_strings.count) {
            return str_duplicate(conv->shared_strings.strings[index]);
        }
        return str_duplicate(value);
    }

    /* Handle boolean */
    if (type_attr && strcmp(type_attr, "b") == 0) {
        int bool_val = atoi(value);
        return str_duplicate(bool_val ? "TRUE" : "FALSE");
    }

    /* Handle inline string */
    if (type_attr && (strcmp(type_attr, "str") == 0 || strcmp(type_attr, "inlineStr") == 0)) {
        return str_duplicate(value);
    }

    /* Handle numeric value with style */
    if (style_attr) {
        int        style_id  = atoi(style_attr);
        formatType ftype     = get_format_type(style_id, &conv->styles);
        double     num_value = atof(value);

        if (ftype == FORMAT_DATE) {
            if (conv->options.dateformat) {
                return format_date(num_value, conv->options.dateformat, conv->workbook.date1904);
            } else {
                return format_date(num_value, NULL, conv->workbook.date1904);
            }
        } else if (ftype == FORMAT_TIME) {
            if (conv->options.timeformat) {
                return format_time(num_value, conv->options.timeformat);
            } else {
                return format_time(num_value, NULL);
            }
        } else if (ftype == FORMAT_FLOAT) {
            /* Python xlsx2csv logic: floatformat only applies in specific cases:
             * 1. If original value contains 'e' or 'E' (scientific notation) - always apply
             * floatformat
             * 2. If format string starts with '0.0' (like '0.00', '0.0000') - apply floatformat if
             * specified
             * 3. If format is 'general' - don't apply floatformat (unless scientific notation)
             */
            const char *format_str = NULL;
            for (int i = 0; i < conv->styles.format_count; i++) {
                if (conv->styles.formats[i].id == style_id) {
                    format_str = conv->styles.formats[i].format_code;
                    break;
                }
            }

            bool format_starts_with_0_0 = format_str && strncmp(format_str, "0.0", 3) == 0;
            bool has_scientific         = strchr(value, 'e') != NULL || strchr(value, 'E') != NULL;

            /* Apply floatformat only if:
             * - original value has scientific notation (always), OR
             * - format starts with '0.0' AND floatformat is specified
             * For 'general' format without scientific notation, use default formatting
             */
            /* Check if original value string contains negative zero */
            double parsed_value = atof(value);
            bool   is_negative_zero =
                (strcmp(value, "-0") == 0) ||
                (value[0] == '-' && fabs(parsed_value) < 1e-300 && signbit(parsed_value));
            if (has_scientific && conv->options.floatformat) {
                /* Scientific notation: always apply floatformat if specified */
                char *result =
                    format_float(num_value, conv->options.floatformat, conv->options.scifloat);
                /* Preserve negative zero if original was negative zero */
                if (is_negative_zero && strcmp(result, "0") == 0) {
                    free(result);
                    return str_duplicate("-0");
                }
                return result;
            } else if (format_starts_with_0_0 && conv->options.floatformat) {
                /* Format starts with '0.0': apply floatformat if specified */
                char *result =
                    format_float(num_value, conv->options.floatformat, conv->options.scifloat);
                /* Preserve negative zero if original was negative zero */
                if (is_negative_zero && strcmp(result, "0") == 0) {
                    free(result);
                    return str_duplicate("-0");
                }
                return result;
            } else {
                /* Default formatting (no floatformat) */
                return format_float(num_value, NULL, conv->options.scifloat);
            }
        }
    }

    /* Default numeric handling */
    if (type_attr && strcmp(type_attr, "n") == 0) {
        double num_value = atof(value);

        /* Check if original value contains scientific notation */
        bool has_scientific = strchr(value, 'e') != NULL || strchr(value, 'E') != NULL;

        /* Check if value is integer (or very close to integer) */
        bool is_integer = fabs(num_value - round(num_value)) < 1e-10;

        /* Check if original value is negative zero */
        bool is_negative_zero =
            (strcmp(value, "-0") == 0) || (value[0] == '-' && fabs(num_value) < 1e-300);

        /* Apply floatformat only if original value has scientific notation */
        if (conv->options.floatformat && has_scientific) {
            return format_float(num_value, conv->options.floatformat, conv->options.scifloat);
        }

        /* Apply scifloat formatting if enabled (but not for integers) */
        if (conv->options.scifloat && !is_integer) {
            return format_float(num_value, NULL, conv->options.scifloat);
        }

        /* Otherwise use default formatting */
        char buffer[256];

        if (is_integer) {
            /* Integer: display as integer (no decimal point) */
            if (is_negative_zero && conv->options.floatformat) {
                /* When floatformat is specified, preserve negative zero */
                snprintf(buffer, sizeof(buffer), "-0");
            } else if (fabs(num_value) < 1e-10) {
                /* Python xlsx2csv: "%i" % Decimal(repr(float("-0"))) -> "0" */
                snprintf(buffer, sizeof(buffer), "0");
            } else {
                snprintf(buffer, sizeof(buffer), "%.0f", num_value);
            }
        } else if (strchr(value, 'e') || strchr(value, 'E')) {
            /* Original was in scientific notation, convert to decimal format with %f */
            snprintf(buffer, sizeof(buffer), "%f", num_value);

            /* Strip trailing zeros */
            char *p = strchr(buffer, '.');
            if (p) {
                char *end = buffer + strlen(buffer) - 1;
                while (end > p && *end == '0') {
                    *end = '\0';
                    end--;
                }
                if (*end == '.') {
                    *end = '\0';
                }
            }
        } else {
            /* Use %f (6 decimal places by default), then strip trailing zeros
             * This matches Python's behavior: ("%f" % data).rstrip('0').rstrip('.')
             */
            snprintf(buffer, sizeof(buffer), "%f", num_value);

            /* Strip trailing zeros */
            char *p = strchr(buffer, '.');
            if (p) {
                char *end = buffer + strlen(buffer) - 1;
                while (end > p && *end == '0') {
                    *end = '\0';
                    end--;
                }
                if (*end == '.') {
                    *end = '\0';
                }
            }
        }

        return str_duplicate(buffer);
    }

    /* Default: return as-is */
    return str_duplicate(value);
}
