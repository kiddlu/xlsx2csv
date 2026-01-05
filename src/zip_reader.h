#ifndef _ZIP_READER_H
#define _ZIP_READER_H

#include <stddef.h>

/* ZIP file operations */
void *zip_open_file(const char *filename);
void *zip_open_stdin(void);
void  xlsx_zip_close(void *handle);

/* ZIP entry operations */
void *zip_file_open(void *zip_handle, const char *filename);
int   zip_file_read(void *file_handle, void *buffer, size_t size);
void  zip_file_close(void *file_handle);

/* Utility functions */
char *zip_read_file_to_string(void *zip_handle, const char *filename);

#endif /* _ZIP_READER_H */
