
#include "pch.h" // use stdafx.h in Visual Studio 2017 and earlier

#include <stdlib.h>
#include <string.h>
#include <stdint.h>
#include <stdbool.h>
#include <stdio.h>

#include <Windows.h>
#include < combaseapi.h >
#include <tlhelp32.h>
#include <tchar.h>


#include "proc_list.h"

# define MAX_NPROC 1024

int new_app_list(app_list_s* list) {
    list->count = 0;
    list->names = (char**)malloc(sizeof(char*) * MAX_NPROC);
    if (list->names == NULL) { return EXIT_FAILURE; }
    for (int i = 0; i < MAX_NPROC; i++) {
        list->names[i] = (char*)malloc(sizeof(char) * MAX_PATH);
        if (list->names[i] == NULL) { return EXIT_FAILURE; }
    }
    return EXIT_SUCCESS;
}

int del_app_list(app_list_s* list) {
    for (int i = 0; i < MAX_NPROC; i++) {
        free(list->names[i]);
    }
    free(list->names);
    list->names = NULL;
    return EXIT_SUCCESS;
}

int get_process_list(app_list_s* process_list) {
    if (process_list->names == NULL) { return -1; }

    HANDLE hSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0);
    if (hSnapShot == INVALID_HANDLE_VALUE) {
        return -1;
    }

    PROCESSENTRY32 pe;
    pe.dwSize = sizeof(pe);

    if (!Process32First(hSnapShot, &pe)) {
        CloseHandle(hSnapShot);    return -1;
    }

    int i = 0;
    do {
        size_t rv;
        wcstombs_s(&rv, process_list->names[i], MAX_PATH, pe.szExeFile, _TRUNCATE);
        ++i;
    } while (Process32Next(hSnapShot, &pe) && i < MAX_NPROC);

    process_list->count = i;

    CloseHandle(hSnapShot);

    return EXIT_SUCCESS;
}
