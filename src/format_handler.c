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
    {"0.0%",                  FORMAT_PERCENTAGE},
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

    /* Check for date/time pattern (case-insensitive)
     * Time-only formats: contain h/H or s/S but NOT y/Y or d/D
     * Date formats: contain y/Y or d/D
     * DateTime formats: contain both date and time components
     */
    bool has_year  = strstr(format_str, "y") || strstr(format_str, "Y");
    bool has_month = strstr(format_str, "m") || strstr(format_str, "M");
    bool has_day   = strstr(format_str, "d") || strstr(format_str, "D");
    bool has_hour  = strstr(format_str, "h") || strstr(format_str, "H");
    bool has_sec   = strstr(format_str, "s") || strstr(format_str, "S");

    bool has_date_component = has_year || has_day;
    bool has_time_component = has_hour || has_sec;

    if (has_time_component && !has_date_component) {
        /* Time-only format: HH:MM:SS, h:mm, etc. */
        return FORMAT_TIME;
    } else if (has_date_component || has_month) {
        /* Date or DateTime format */
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

    /* Check if format contains both date and time components
     * In Excel formats:
     * - H/h = hours (time component)
     * - S/s = seconds (time component)
     * - m/M near h/H = minutes (time component)
     * - m/M alone = months (date component)
     */
    bool has_time_component = false;
    if (format) {
        /* Check for hours or seconds (clear time indicators) */
        has_time_component = strstr(format, "H") || strstr(format, "h") || strstr(format, "S") ||
                             strstr(format, "s");
        /* But only if it also has date components */
        bool has_date_component = strstr(format, "Y") || strstr(format, "y") ||
                                  strstr(format, "D") || strstr(format, "d");
        has_time_component = has_time_component && has_date_component;
    }

    /* Convert Excel serial date to year, month, day
     * Excel uses 1900-01-01 as day 1, with a leap year bug
     * Python xlsx2csv uses 1899-12-30 as day 0 (day 1 = 1899-12-31)
     * We'll use Python's epoch (1899-12-30) to match its behavior
     */
    int total_days;
    if (date1904) {
        /* 1904-01-01 is day 0 in 1904 system
         * Convert to days since 1899-12-30
         */
        total_days = (int)value + 1462; /* Days from 1899-12-30 to 1904-01-01 */
    } else {
        /* Excel day 1 = 1900-01-01, but we use 1899-12-30 as epoch
         * So Excel day 1 = our day 2
         * But Python xlsx2csv outputs day 1 as 1899-12-31, so day 1 = epoch + 1
         * This suggests Python uses: Excel day N = epoch + N
         */
        total_days = (int)value;
    }

    /* Use standard C library for date calculation */
    /* 1899-12-30 = Unix epoch - 25569 days */
    /* But it's easier to calculate manually */

    /* Start from 1899-12-30 and add total_days */
    int year  = 1899;
    int month = 12;
    int day   = 30;

    /* Add days */
    day += total_days;

    /* Days in each month */
    int days_in_month[] = {31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31};

    /* Advance through years and months */
    while (true) {
        /* Check for leap year */
        bool is_leap = (year % 4 == 0 && year % 100 != 0) || (year % 400 == 0);
        if (is_leap) {
            days_in_month[1] = 29;
        } else {
            days_in_month[1] = 28;
        }

        int days_in_current_month = days_in_month[month - 1];

        if (day <= days_in_current_month) {
            break;
        }

        day -= days_in_current_month;
        month++;

        if (month > 12) {
            month = 1;
            year++;
        }
    }

    char buffer[256];

    if (has_time_component) {
        /* Format as DateTime: YYYY-MM-DD HH:MM:SS */
        /* Extract time from fractional part */
        double time_fraction = value - (int)value;
        int    total_seconds = (int)(time_fraction * 86400);
        int    hours         = total_seconds / 3600;
        int    minutes       = (total_seconds % 3600) / 60;
        int    seconds       = total_seconds % 60;

        snprintf(buffer,
                 sizeof(buffer),
                 "%04d-%02d-%02d %02d:%02d:%02d",
                 year,
                 month,
                 day,
                 hours,
                 minutes,
                 seconds);
    } else {
        /* Use default date format: YYYY-MM-DD */
        (void)format; /* Suppress unused parameter warning for non-datetime case */
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
char *format_float(double value, const char *format, bool scifloat, const char *original_str)
{
    char buffer[256];

    if (format) {
        /* Check if value is integer AND original string doesn't indicate float
         * True integers: value is close to integer AND no scientific notation in original
         * Float zeros: value is near-zero BUT from scientific notation (e.g., 1e-100)
         */
        bool is_integer_value = fabs(value - round(value)) < 1e-10;
        bool has_scientific =
            original_str && (strchr(original_str, 'e') || strchr(original_str, 'E'));

        if (is_integer_value && !has_scientific) {
            /* Python behavior: true integers are output as integers even with --floatformat */
            snprintf(buffer, sizeof(buffer), "%.0f", value);
        } else {
            /* Float values or scientific notation zeros: apply format */
#pragma GCC diagnostic push
#pragma GCC diagnostic ignored "-Wformat-nonliteral"
            snprintf(buffer, sizeof(buffer), format, value);
#pragma GCC diagnostic pop
        }

        return str_duplicate(buffer);
    } else if (scifloat) {
        /* Python xlsx2csv's --sci-float behavior:
         * Use regular decimal format (not scientific notation)
         * But for integer values (not from scientific notation), output as integer
         */

        /* Check if original value contains scientific notation */
        bool has_scientific =
            original_str && (strchr(original_str, 'e') || strchr(original_str, 'E'));

        /* Check if value is integer */
        bool is_integer_value = fabs(value - round(value)) < 1e-10;

        if (is_integer_value && !has_scientific) {
            /* Integer from normal notation: output as integer even with --sci-float */
            snprintf(buffer, sizeof(buffer), "%.0f", value);
        } else {
            /* Float or from scientific notation: use %f format */
            snprintf(buffer, sizeof(buffer), "%f", value);

            /* Strip trailing zeros for non-zero values (but not for values from scientific notation
             * that became zero) */
            double result_value = atof(buffer);
            bool   is_zero      = fabs(result_value) < 1e-300;

            /* Only strip trailing zeros if:
             * 1. Value is not zero, OR
             * 2. Value is zero but NOT from scientific notation
             */
            if (!is_zero) {
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
            /* For values from scientific notation that became zero, keep "0.000000" or "-0.000000"
             */
        }
    } else if (fabs(value - round(value)) < 1e-9 && fabs(value) > 1e-200) {
        /* Integer value (but not subnormal/tiny values)
         * Exclude subnormal numbers (< 1e-200) to match Python behavior
         * which formats them as "0.000000" not "0"
         */
        snprintf(buffer, sizeof(buffer), "%.0f", value);
    } else if (fabs(value) < 1e-200 && value != 0.0) {
        /* Subnormal/tiny values: use %f to get "0.000000" or "-0.000000"
         * This matches Python's behavior for extremely small values
         */
        snprintf(buffer, sizeof(buffer), "%f", value);
        /* Do NOT strip trailing zeros for tiny values - Python keeps them */
    } else if (fabs(value) < 1e-10) {
        /* Value is effectively zero (not subnormal, just zero) */
        if (value < 0.0 || signbit(value)) {
            snprintf(buffer, sizeof(buffer), "-0");
        } else {
            snprintf(buffer, sizeof(buffer), "0");
        }
    } else {
        /* Default: Use %.15g format to preserve precision
         * This matches Python xlsx2csv's default behavior
         */
        snprintf(buffer, sizeof(buffer), "%.15g", value);
    }

    return str_duplicate(buffer);
}

/* Apply Excel number format to a value
 * Returns formatted string, or NULL if format not supported
 * Caller must free the returned string
 */
static char *apply_excel_format(double value, const char *format_code)
{
    if (!format_code) {
        return NULL;
    }

    /* Python xlsx2csv only applies simple formats without # or ,
     * Only handle exact matches: "0.00", "0", "0.00E+00"
     */

    /* Handle format "0.00" - round to 2 decimal places */
    if (strcmp(format_code, "0.00") == 0) {
        char buffer[256];
        snprintf(buffer, sizeof(buffer), "%.2f", value);
        return str_duplicate(buffer);
    }

    /* Handle format "0" - no decimal places */
    if (strcmp(format_code, "0") == 0) {
        char buffer[256];
        snprintf(buffer, sizeof(buffer), "%.0f", value);
        return str_duplicate(buffer);
    }

    /* Handle scientific notation format "0.00E+00"
     * Python outputs with 6 decimal places (%.6f)
     */
    if (strcmp(format_code, "0.00E+00") == 0 || strcmp(format_code, "0.00e+00") == 0) {
        char buffer[256];
        snprintf(buffer, sizeof(buffer), "%.6f", value);
        return str_duplicate(buffer);
    }

    /* All other formats (including #,##0, #,##0.00, etc.) are not applied
     * Return NULL to use default formatting
     */
    return NULL;
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

    /* Handle Excel error values - preserve them as-is
     * Common Excel errors: #DIV/0!, #N/A, #NAME?, #NULL!, #NUM!, #REF!, #VALUE!
     */
    if (value[0] == '#') {
        return str_duplicate(value);
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
            /* Get the format string for datetime detection */
            int         fmt_id     = -1;
            const char *format_str = NULL;
            if (style_id >= 0 && style_id < conv->styles.cell_xfs_count) {
                fmt_id = conv->styles.cell_xfs[style_id];
            }
            for (int i = 0; i < conv->styles.format_count; i++) {
                if (conv->styles.formats[i].id == fmt_id) {
                    format_str = conv->styles.formats[i].format_code;
                    break;
                }
            }

            if (conv->options.dateformat) {
                return format_date(num_value, conv->options.dateformat, conv->workbook.date1904);
            } else {
                /* Pass Excel format string to detect DateTime vs Date-only */
                return format_date(num_value, format_str, conv->workbook.date1904);
            }
        } else if (ftype == FORMAT_TIME) {
            if (conv->options.timeformat) {
                return format_time(num_value, conv->options.timeformat);
            } else {
                return format_time(num_value, NULL);
            }
        } else if (ftype == FORMAT_PERCENTAGE) {
            /* Percentage values are stored as decimals in Excel (0.5 = 50%)
             * Python xlsx2csv applies floatformat ONLY if the percentage format starts with "0.0"
             * e.g., "0.0%" applies floatformat, but "0%" does not
             */
            /* Get the format string for this style */
            int         fmt_id     = -1;
            const char *format_str = NULL;
            if (style_id >= 0 && style_id < conv->styles.cell_xfs_count) {
                fmt_id = conv->styles.cell_xfs[style_id];
            }
            for (int i = 0; i < conv->styles.format_count; i++) {
                if (conv->styles.formats[i].id == fmt_id) {
                    format_str = conv->styles.formats[i].format_code;
                    break;
                }
            }

            bool format_starts_with_0_0 = format_str && strncmp(format_str, "0.0", 3) == 0;
            if (format_starts_with_0_0 && conv->options.floatformat) {
                return format_float(
                    num_value, conv->options.floatformat, conv->options.scifloat, value);
            } else {
                return format_float(num_value, NULL, conv->options.scifloat, value);
            }
        } else if (ftype == FORMAT_FLOAT) {
            /* Get the format ID for this style */
            int fmt_id = -1;
            if (style_id >= 0 && style_id < conv->styles.cell_xfs_count) {
                fmt_id = conv->styles.cell_xfs[style_id];
            }

            /* Get the format string for this format ID */
            const char *format_str = NULL;
            for (int i = 0; i < conv->styles.format_count; i++) {
                if (conv->styles.formats[i].id == fmt_id) {
                    format_str = conv->styles.formats[i].format_code;
                    break;
                }
            }

            /* If no custom format found, check standard formats */
            if (!format_str && fmt_id >= 0) {
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
                    {37, "#,##0 ;(#,##0)"          },
                    {38, "#,##0 ;[red](#,##0)"     },
                    {39, "#,##0.00;(#,##0.00)"     },
                    {40, "#,##0.00;[red](#,##0.00)"},
                    {48, "##0.0e+0"                }
                };
                for (size_t i = 0; i < sizeof(standard_formats) / sizeof(standard_formats[0]);
                     i++) {
                    if (standard_formats[i].id == fmt_id) {
                        format_str = standard_formats[i].format;
                        break;
                    }
                }
            }

            /* Python behavior is complex:
             * - Standard formats (like "0.00", "#,##0.00", etc.) are NOT applied when --floatformat
             * exists
             * - Custom formats (like "0.00_ ", "0.00 ") ARE applied (keep Excel precision)
             * Check format string patterns
             */
            bool is_standard_format = false;
            if (format_str) {
                /* Standard formats: starts with # or 0, no decorators like _ or spaces after
                 * numbers */
                if (strcmp(format_str, "0") == 0 || strcmp(format_str, "0.00") == 0 ||
                    strcmp(format_str, "0.00E+00") == 0 || strcmp(format_str, "0.00e+00") == 0 ||
                    strncmp(format_str, "#,##0", 5) == 0) {
                    is_standard_format = true;
                }
            }

            bool has_custom_format =
                format_str && !is_standard_format && strncmp(format_str, "0.0", 3) == 0;

            /* Apply Excel format in these cases:
             * 1. No --floatformat: apply any format starting with "0.0"
             * 2. Has --floatformat WITH custom format: DO NOT apply Excel format here
             *    (will be handled in floatformat path with custom stripping logic)
             */
            bool should_apply_excel = false;
            if (!conv->options.floatformat && format_str && strncmp(format_str, "0.0", 3) == 0) {
                should_apply_excel = true;
            }
            /* Note: When floatformat is present, skip Excel format application
             * and let the floatformat path handle it */

            if (should_apply_excel) {
                /* Custom format: apply Excel format and keep precision */
                char *excel_formatted = apply_excel_format(num_value, format_str);
                if (excel_formatted) {
                    return excel_formatted;
                }
                /* If apply_excel_format returns NULL, handle custom format
                 * Python xlsx2csv precision rule for custom formats:
                 * output_precision = excel_decimal_places + modifier_length
                 * where modifier_length = characters after the last '0' in format
                 * Examples:
                 * - "0.00_ " (2 decimals + 2 chars) -> 4 decimal precision
                 * - "0.00 "  (2 decimals + 1 char)  -> 3 decimal precision
                 * - "0.0000" (4 decimals + 0 chars) -> 4 decimal precision
                 * Then strip trailing zeros for custom formats
                 */
                int         decimal_places = 2;  // default
                const char *dot            = strchr(format_str, '.');
                if (dot) {
                    decimal_places = 0;
                    const char *p  = dot + 1;
                    while (*p == '0') {
                        decimal_places++;
                        p++;
                    }
                    /* Count modifier length (characters after last '0') */
                    int modifier_length = strlen(p);

                    /* For custom formats, add modifier length to precision */
                    int output_decimal_places = decimal_places;
                    if (has_custom_format) {
                        output_decimal_places = decimal_places + modifier_length;
                    }

                    char format_buf[32];
                    snprintf(format_buf, sizeof(format_buf), "%%.%df", output_decimal_places);

                    char buffer[256];
#pragma GCC diagnostic push
#pragma GCC diagnostic ignored "-Wformat-nonliteral"
                    snprintf(buffer, sizeof(buffer), format_buf, num_value);
#pragma GCC diagnostic pop

                    /* Strip trailing zeros for custom formats
                     * Python behavior differs based on --floatformat:
                     * - With --floatformat: always strip trailing zeros
                     * - Without --floatformat: only strip if last digit is NOT 0
                     * Examples:
                     * - With --floatformat: 5.10 -> 5.1
                     * - Without --floatformat: 83.5600 -> 83.5600 (keep), -1.7215 -> -1.7215
                     * (strip)
                     */
                    if (has_custom_format) {
                        bool should_strip = false;
                        if (conv->options.floatformat) {
                            /* With --floatformat: always strip */
                            should_strip = true;
                        } else {
                            /* Without --floatformat: only strip if last digit is not 0 */
                            size_t len = strlen(buffer);
                            if (len > 0 && buffer[len - 1] != '0') {
                                should_strip = true;
                            }
                        }

                        if (should_strip) {
                            char *p_buf = strchr(buffer, '.');
                            if (p_buf) {
                                char *end = buffer + strlen(buffer) - 1;
                                while (end > p_buf && *end == '0') {
                                    *end = '\0';
                                    end--;
                                }
                                if (*end == '.') {
                                    *end = '\0';
                                }
                            }
                        }
                    }

                    return str_duplicate(buffer);
                }
            }

            /* Second priority: Apply floatformat option if specified
             * Python behavior:
             * - For standard formats (#,##0.00, etc.): do NOT apply floatformat
             * - For custom formats (0.00_ , etc.): apply floatformat EVEN for integer values
             */
            /* Check if original value string contains negative zero */
            double parsed_value = atof(value);
            bool   is_negative_zero =
                (strcmp(value, "-0") == 0) ||
                (value[0] == '-' && fabs(parsed_value) < 1e-300 && signbit(parsed_value));

            if (conv->options.floatformat && has_custom_format) {
                /* With --floatformat and custom format: apply floatformat */
                /* For custom formats, even integer values should be formatted with decimals */
                char format_buf[256];
#pragma GCC diagnostic push
#pragma GCC diagnostic ignored "-Wformat-nonliteral"
                snprintf(format_buf, sizeof(format_buf), conv->options.floatformat, num_value);
#pragma GCC diagnostic pop
                return str_duplicate(format_buf);
            } else if (conv->options.floatformat && !is_standard_format) {
                /* With --floatformat but no Excel format: apply for floats only */
                char *result = format_float(
                    num_value, conv->options.floatformat, conv->options.scifloat, value);
                if (is_negative_zero && strcmp(result, "0") == 0) {
                    free(result);
                    return str_duplicate("-0");
                }
                return result;
            } else {
                /* Without --floatformat OR with standard format: default formatting */
                return format_float(num_value, NULL, conv->options.scifloat, value);
            }
        }
    }

    /* Default numeric handling */
    if (type_attr && strcmp(type_attr, "n") == 0) {
        double num_value = atof(value);

        /* Check if original value contains scientific notation */
        bool has_scientific = strchr(value, 'e') != NULL || strchr(value, 'E') != NULL;

        /* Check if value is integer (or very close to integer)
         * BUT: if original value was in scientific notation, treat as float
         */
        bool is_integer = !has_scientific && (fabs(num_value - round(num_value)) < 1e-10);

        /* Check if original value is negative zero */
        bool is_negative_zero =
            (strcmp(value, "-0") == 0) || (value[0] == '-' && fabs(num_value) < 1e-300);

        /* Python xlsx2csv behavior for --floatformat:
         * - Only applies floatformat if the cell has an Excel number format
         * - For cells without Excel format: use default %.15g precision
         * - Exception: scientific notation values always apply floatformat
         */
        if (conv->options.floatformat && has_scientific) {
            /* Scientific notation: always apply floatformat */
            return format_float(
                num_value, conv->options.floatformat, conv->options.scifloat, value);
        }

        /* For cells without Excel format: do NOT apply --floatformat option
         * (This branch is for cells without style or with style but no custom number format)
         * Use default formatting instead */
        if (conv->options.scifloat && !is_integer) {
            return format_float(num_value, NULL, conv->options.scifloat, value);
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
            /* Original was in scientific notation, convert to decimal format with %f
             * Keep the default 6 decimal places to match Python's behavior
             */
            snprintf(buffer, sizeof(buffer), "%f", num_value);
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
