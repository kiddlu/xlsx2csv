#include "xlsx2csv.h"
#include <stdlib.h>
#include <string.h>
#include <strings.h>
#include <zip.h>
#include <errno.h>

/* Open XLSX file (which is a ZIP archive) */
void *zip_open_file(const char *filename) {
    int err = 0;
    zip_t *za = zip_open(filename, ZIP_RDONLY, &err);
    
    if (za == NULL) {
        zip_error_t error;
        zip_error_init_with_code(&error, err);
        fprintf(stderr, "Error opening %s: %s\n", filename, zip_error_strerror(&error));
        zip_error_fini(&error);
        return NULL;
    }
    
    return (void *)za;
}

/* Open from STDIN (read into memory buffer) */
void *zip_open_stdin(void) {
    /* Read all data from stdin into buffer */
    size_t buffer_size = 4096;
    size_t total_read = 0;
    char *buffer = malloc(buffer_size);
    
    if (!buffer) {
        fprintf(stderr, "Error: Out of memory\n");
        return NULL;
    }
    
    while (1) {
        size_t read_size = fread(buffer + total_read, 1, buffer_size - total_read, stdin);
        total_read += read_size;
        
        if (read_size == 0) {
            break;
        }
        
        if (total_read >= buffer_size) {
            buffer_size *= 2;
            char *new_buffer = realloc(buffer, buffer_size);
            if (!new_buffer) {
                free(buffer);
                fprintf(stderr, "Error: Out of memory\n");
                return NULL;
            }
            buffer = new_buffer;
        }
    }
    
    /* Open ZIP from memory buffer */
    zip_error_t error;
    zip_source_t *src = zip_source_buffer_create(buffer, total_read, 1, &error);
    if (src == NULL) {
        fprintf(stderr, "Error creating zip source: %s\n", zip_error_strerror(&error));
        free(buffer);
        return NULL;
    }
    
    zip_t *za = zip_open_from_source(src, ZIP_RDONLY, &error);
    if (za == NULL) {
        fprintf(stderr, "Error opening zip from stdin: %s\n", zip_error_strerror(&error));
        zip_source_free(src);
        return NULL;
    }
    
    return (void *)za;
}

/* Close ZIP archive */
void xlsx_zip_close(void *handle) {
    if (handle) {
        zip_close((zip_t *)handle);
    }
}

/* Open file within ZIP archive (case-insensitive) */
void *zip_file_open(void *zip_handle, const char *filename) {
    if (!zip_handle || !filename) {
        return NULL;
    }
    
    zip_t *za = (zip_t *)zip_handle;
    
    /* Remove leading slash if present */
    const char *search_name = filename;
    if (search_name[0] == '/') {
        search_name++;
    }
    
    /* Try to find file (case-insensitive) */
    zip_int64_t num_entries = zip_get_num_entries(za, 0);
    for (zip_int64_t i = 0; i < num_entries; i++) {
        const char *name = zip_get_name(za, i, 0);
        if (name && strcasecmp(name, search_name) == 0) {
            zip_file_t *zf = zip_fopen_index(za, i, 0);
            return (void *)zf;
        }
    }
    
    return NULL;
}

/* Read from file within ZIP */
int zip_file_read(void *file_handle, void *buffer, size_t size) {
    if (!file_handle || !buffer) {
        return -1;
    }
    
    zip_file_t *zf = (zip_file_t *)file_handle;
    return (int)zip_fread(zf, buffer, size);
}

/* Close file within ZIP */
void zip_file_close(void *file_handle) {
    if (file_handle) {
        zip_fclose((zip_file_t *)file_handle);
    }
}

/* Read entire file from ZIP to string */
char *zip_read_file_to_string(void *zip_handle, const char *filename) {
    void *file = zip_file_open(zip_handle, filename);
    if (!file) {
        return NULL;
    }
    
    /* Read file in chunks */
    size_t buffer_size = 4096;
    size_t total_read = 0;
    char *buffer = malloc(buffer_size);
    
    if (!buffer) {
        zip_file_close(file);
        return NULL;
    }
    
    while (1) {
        int read_size = zip_file_read(file, buffer + total_read, buffer_size - total_read - 1);
        if (read_size <= 0) {
            break;
        }
        
        total_read += read_size;
        
        if (total_read >= buffer_size - 1) {
            buffer_size *= 2;
            char *new_buffer = realloc(buffer, buffer_size);
            if (!new_buffer) {
                free(buffer);
                zip_file_close(file);
                return NULL;
            }
            buffer = new_buffer;
        }
    }
    
    buffer[total_read] = '\0';
    zip_file_close(file);
    
    return buffer;
}

