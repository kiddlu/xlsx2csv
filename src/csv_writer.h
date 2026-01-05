#ifndef _CSV_WRITER_H
#define _CSV_WRITER_H

#include <stdio.h>

#include "xlsx2csv.h"

/* Forward declaration */
typedef struct csvWriter csvWriter;

/* CSV Writer functions */
csvWriter *csv_writer_create(FILE *fp, xlsxOptions *options);
void       csv_writer_free(csvWriter *writer);
int        csv_write_row(csvWriter *writer, char **fields, int field_count);
int        csv_write_field(csvWriter *writer, const char *field);
void       csv_writer_reset_row(csvWriter *writer);
void       csv_writer_set_field_count(csvWriter *writer, int count);

#endif /* _CSV_WRITER_H */
