#ifndef _FORMAT_HANDLER_H
#define _FORMAT_HANDLER_H

#include <stdbool.h>

#include "xlsx2csv.h"

/* Format handler functions */
char      *format_cell_value(const char        *value,
                             const char        *type_attr,
                             const char        *style_attr,
                             xlsx2csvConverter *conv);
formatType get_format_type(int style_id, styleInfo *styles);
char      *format_date(double value, const char *format, bool date1904);
char      *format_time(double value, const char *format);
char      *format_float(double value, const char *format, bool scifloat, const char *original_str);

#endif /* _FORMAT_HANDLER_H */
