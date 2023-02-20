#pragma once

#ifdef PROCLIST_EXPORTS
#define PROC_LIST_API __declspec(dllexport)
#else
#define PROC_LIST_API __declspec(dllimport)
#endif


typedef struct APP_LIST {
    size_t count;
    char** names;
} app_list_s;

PROC_LIST_API int new_app_list(app_list_s* list);
PROC_LIST_API int del_app_list(app_list_s* list);
PROC_LIST_API int get_process_list(app_list_s* process_list);