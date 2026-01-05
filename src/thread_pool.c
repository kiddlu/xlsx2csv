/* Standard library headers */
#include <pthread.h>
#include <stdlib.h>
#include <string.h>

/* Project headers */
#include "thread_pool.h"

/* Result structure for ordered output */
typedef struct {
    int   task_id;
    void *result;
} result_t;

/* Thread pool structure */
struct thread_pool {
    pthread_t      *threads;
    int             num_threads;
    task_t         *task_queue;
    int             queue_size;
    int             queue_head;
    int             queue_tail;
    int             queue_count;
    result_t       *results;
    int             results_size;
    int             results_count;
    pthread_mutex_t mutex;
    pthread_cond_t  cond;
    pthread_cond_t  result_cond;
    bool            shutdown;
};

/* Worker thread function */
static void *worker_thread(void *arg)
{
    thread_pool_t *pool = (thread_pool_t *)arg;

    while (1) {
        pthread_mutex_lock(&pool->mutex);

        /* Wait for tasks or shutdown */
        while (pool->queue_count == 0 && !pool->shutdown) {
            pthread_cond_wait(&pool->cond, &pool->mutex);
        }

        if (pool->shutdown) {
            pthread_mutex_unlock(&pool->mutex);
            break;
        }

        /* Get task from queue */
        task_t task      = pool->task_queue[pool->queue_head];
        pool->queue_head = (pool->queue_head + 1) % pool->queue_size;
        pool->queue_count--;

        pthread_mutex_unlock(&pool->mutex);

        /* Execute task */
        task.func(task.arg);

        /* Store result */
        pthread_mutex_lock(&pool->mutex);
        if (pool->results_count < pool->results_size) {
            pool->results[pool->results_count].task_id = task.task_id;
            pool->results[pool->results_count].result  = task.arg;
            pool->results_count++;
        }
        pthread_cond_signal(&pool->result_cond);
        pthread_mutex_unlock(&pool->mutex);
    }

    return NULL;
}

/* Create thread pool */
thread_pool_t *thread_pool_create(int num_threads)
{
    if (num_threads <= 0) {
        return NULL;
    }

    thread_pool_t *pool = calloc(1, sizeof(thread_pool_t));
    if (!pool) {
        return NULL;
    }

    pool->num_threads   = num_threads;
    pool->queue_size    = 1024; /* Reasonable default */
    pool->results_size  = 1024;
    pool->queue_head    = 0;
    pool->queue_tail    = 0;
    pool->queue_count   = 0;
    pool->results_count = 0;
    pool->shutdown      = false;

    pool->threads    = calloc((size_t)num_threads, sizeof(pthread_t));
    pool->task_queue = calloc((size_t)pool->queue_size, sizeof(task_t));
    pool->results    = calloc((size_t)pool->results_size, sizeof(result_t));

    if (!pool->threads || !pool->task_queue || !pool->results) {
        free(pool->threads);
        free(pool->task_queue);
        free(pool->results);
        free(pool);
        return NULL;
    }

    pthread_mutex_init(&pool->mutex, NULL);
    pthread_cond_init(&pool->cond, NULL);
    pthread_cond_init(&pool->result_cond, NULL);

    /* Create worker threads */
    for (int i = 0; i < num_threads; i++) {
        if (pthread_create(&pool->threads[i], NULL, worker_thread, pool) != 0) {
            /* Cleanup on failure */
            pool->shutdown = true;
            pthread_cond_broadcast(&pool->cond);
            for (int j = 0; j < i; j++) {
                pthread_join(pool->threads[j], NULL);
            }
            thread_pool_destroy(pool);
            return NULL;
        }
    }

    return pool;
}

/* Destroy thread pool */
void thread_pool_destroy(thread_pool_t *pool)
{
    if (!pool) {
        return;
    }

    /* Signal shutdown */
    pthread_mutex_lock(&pool->mutex);
    pool->shutdown = true;
    pthread_cond_broadcast(&pool->cond);
    pthread_mutex_unlock(&pool->mutex);

    /* Wait for all threads */
    for (int i = 0; i < pool->num_threads; i++) {
        pthread_join(pool->threads[i], NULL);
    }

    /* Cleanup */
    pthread_mutex_destroy(&pool->mutex);
    pthread_cond_destroy(&pool->cond);
    pthread_cond_destroy(&pool->result_cond);
    free(pool->threads);
    free(pool->task_queue);
    free(pool->results);
    free(pool);
}

/* Submit task to thread pool */
int thread_pool_submit(thread_pool_t *pool, task_func_t func, void *arg, int task_id)
{
    if (!pool || !func) {
        return -1;
    }

    pthread_mutex_lock(&pool->mutex);

    /* Check if queue is full */
    if (pool->queue_count >= pool->queue_size) {
        pthread_mutex_unlock(&pool->mutex);
        return -1; /* Queue full */
    }

    /* Add task to queue */
    pool->task_queue[pool->queue_tail].func    = func;
    pool->task_queue[pool->queue_tail].arg     = arg;
    pool->task_queue[pool->queue_tail].task_id = task_id;
    pool->queue_tail                           = (pool->queue_tail + 1) % pool->queue_size;
    pool->queue_count++;

    pthread_cond_signal(&pool->cond);
    pthread_mutex_unlock(&pool->mutex);

    return 0;
}

/* Wait for all tasks to complete */
void thread_pool_wait(thread_pool_t *pool)
{
    if (!pool) {
        return;
    }

    pthread_mutex_lock(&pool->mutex);
    /* Wait until queue is empty (all tasks have been taken by workers) */
    while (pool->queue_count > 0) {
        pthread_cond_wait(&pool->result_cond, &pool->mutex);
    }
    pthread_mutex_unlock(&pool->mutex);

    /* Note: We don't wait for results_count here because results are added
     * asynchronously. The caller should check results_count separately if needed. */
}

/* Get result count */
int thread_pool_get_result_count(thread_pool_t *pool)
{
    if (!pool) {
        return 0;
    }
    return pool->results_count;
}

/* Get result by index */
void *thread_pool_get_result(thread_pool_t *pool, int index)
{
    if (!pool || index < 0 || index >= pool->results_count) {
        return NULL;
    }
    return pool->results[index].result;
}
