#ifndef _XML_PARSER_H
#define _XML_PARSER_H

#include <stdio.h>

#include "xlsx2csv.h"

/* XML parser functions */
int parse_content_types(xlsx2csvConverter *conv);
int parse_workbook(xlsx2csvConverter *conv);
int parse_shared_strings(xlsx2csvConverter *conv);
int parse_styles(xlsx2csvConverter *conv);
int parse_worksheet(xlsx2csvConverter *conv, int sheet_index, FILE *outfile);

#endif /* _XML_PARSER_H */
