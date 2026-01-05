#ifndef _UTILS_H
#define _UTILS_H

#include <stdbool.h>

/* String utilities */
char *str_duplicate(const char *str);
char *str_concat(const char *str1, const char *str2);
bool  str_match_pattern(const char *str, const char *pattern);

/* Excel column utilities */
int   column_name_to_index(const char *column);
char *column_index_to_name(int index);

/* Validation utilities */
bool is_numeric(const char *str);

#endif /* _UTILS_H */
