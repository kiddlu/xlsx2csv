/* Standard library headers */
#include <ctype.h>
#include <regex.h>
#include <stdlib.h>
#include <string.h>

/* Project headers */
#include "utils.h"

/* Duplicate string */
char *str_duplicate(const char *str)
{
    if (!str) {
        return NULL;
    }

    size_t len = strlen(str);
    char  *dup = malloc(len + 1);
    if (dup) {
        memcpy(dup, str, len + 1);
    }
    return dup;
}

/* Concatenate two strings */
char *str_concat(const char *str1, const char *str2)
{
    if (!str1 || !str2) {
        return NULL;
    }

    size_t len1   = strlen(str1);
    size_t len2   = strlen(str2);
    char  *result = malloc(len1 + len2 + 1);

    if (result) {
        memcpy(result, str1, len1);
        memcpy(result + len1, str2, len2 + 1);
    }

    return result;
}

/* Match string against regex pattern */
bool str_match_pattern(const char *str, const char *pattern)
{
    if (!str || !pattern) {
        return false;
    }

    regex_t regex;
    int     ret = regcomp(&regex, pattern, REG_EXTENDED | REG_NOSUB);
    if (ret != 0) {
        return false;
    }

    ret = regexec(&regex, str, 0, NULL, 0);
    regfree(&regex);

    return (ret == 0);
}

/* Convert column name (A, B, ..., Z, AA, AB, ...) to index (0, 1, ...) */
int column_name_to_index(const char *column)
{
    if (!column) {
        return -1;
    }

    int index = 0;
    for (int i = 0; column[i] != '\0'; i++) {
        if (!isalpha(column[i])) {
            return -1;
        }
        index = index * 26 + (toupper(column[i]) - 'A' + 1);
    }

    return index - 1;
}

/* Convert column index to name */
char *column_index_to_name(int index)
{
    if (index < 0) {
        return NULL;
    }

    char buffer[10];
    int  pos = 0;
    int  col = index + 1;

    while (col > 0) {
        int remainder = (col - 1) % 26;
        buffer[pos++] = 'A' + remainder;
        col           = (col - 1) / 26;
    }

    /* Reverse the string */
    char *result = malloc((size_t)(pos + 1));
    if (result) {
        for (int i = 0; i < pos; i++) {
            result[i] = buffer[pos - 1 - i];
        }
        result[pos] = '\0';
    }

    return result;
}

/* Check if string is numeric */
bool is_numeric(const char *str)
{
    if (!str || *str == '\0') {
        return false;
    }

    int i = 0;

    /* Skip leading whitespace */
    while (isspace(str[i])) {
        i++;
    }

    /* Check for sign */
    if (str[i] == '+' || str[i] == '-') {
        i++;
    }

    bool has_digit = false;
    bool has_dot   = false;
    bool has_exp   = false;

    while (str[i] != '\0') {
        if (isdigit(str[i])) {
            has_digit = true;
        } else if (str[i] == '.' && !has_dot && !has_exp) {
            has_dot = true;
        } else if ((str[i] == 'e' || str[i] == 'E') && !has_exp && has_digit) {
            has_exp   = true;
            has_digit = false; /* Need at least one digit after 'e' */
            i++;
            if (str[i] == '+' || str[i] == '-') {
                i++;
            }
            continue;
        } else if (isspace(str[i])) {
            /* Trailing whitespace is ok */
            break;
        } else {
            return false;
        }
        i++;
    }

    /* Skip trailing whitespace */
    while (isspace(str[i])) {
        i++;
    }

    return has_digit && str[i] == '\0';
}
