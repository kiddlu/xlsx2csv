#ifndef _THREAD_POOL_H
#define _THREAD_POOL_H

#include <stdbool.h>

/* Forward declarations */
typedef struct thread_pool thread_pool_t;

/* Task function type */
typedef void (*task_func_t)(void *arg);

/* Task structure */
typedef struct {
    task_func_t func;
    void       *arg;
    int         task_id;
} task_t;

/* Thread pool API */
thread_pool_t *thread_pool_create(int num_threads);
void           thread_pool_destroy(thread_pool_t *pool);
int            thread_pool_submit(thread_pool_t *pool, task_func_t func, void *arg, int task_id);
void           thread_pool_wait(thread_pool_t *pool);
int            thread_pool_get_result_count(thread_pool_t *pool);
void          *thread_pool_get_result(thread_pool_t *pool, int index);

#endif /* _THREAD_POOL_H */
