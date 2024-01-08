#include <iostream>
#include <Windows.h>
#include <taskschd.h>
#include <comutil.h>
#include <cstringt.h>

#pragma comment(lib, "taskschd.lib")
#pragma comment(lib, "comsupp.lib")

int main()
{
    // 初始化 COM 库
    HRESULT hr = CoInitializeEx(NULL, COINIT_MULTITHREADED);
    if (FAILED(hr))
    {
        std::cout << "Failed to initialize COM library. Error code = " << hr << std::endl;
        return 1;
    }

    // 创建 TaskService 对象
    ITaskService* pService = NULL;
    hr = CoCreateInstance(CLSID_TaskScheduler, NULL, CLSCTX_INPROC_SERVER, IID_ITaskService, (void**)&pService);
    if (FAILED(hr))
    {
        std::cout << "Failed to create TaskService instance. Error code = " << hr << std::endl;
        CoUninitialize();
        return 1;
    }

    // 连接到 Task Scheduler 服务
    hr = pService->Connect(_variant_t(), _variant_t(), _variant_t(), _variant_t());
    if (FAILED(hr))
    {
        std::cout << "Failed to connect to Task Scheduler service. Error code = " << hr << std::endl;
        pService->Release();
        CoUninitialize();
        return 1;
    }

    // 获取 TaskFolder 对象
    ITaskFolder* pRootFolder = NULL;
    hr = pService->GetFolder(_bstr_t(L"\\"), &pRootFolder);
    if (FAILED(hr))
    {
        std::cout << "Failed to get the root task folder. Error code = " << hr << std::endl;
        pService->Release();
        CoUninitialize();
        return 1;
    }

    // 创建一个新的 TaskDefinition 对象
    ITaskDefinition* pTask = NULL;
    hr = pService->NewTask(0, &pTask);
    if (FAILED(hr))
    {
        std::cout << "Failed to create a new task definition. Error code = " << hr << std::endl;
        pRootFolder->Release();
        pService->Release();
        CoUninitialize();
        return 1;
    }

    // 设置任务的基本属性
    ITaskSettings* pSettings = NULL;
    hr = pTask->get_Settings(&pSettings);

    hr = pSettings->put_StartWhenAvailable(VARIANT_TRUE);
    pSettings->Release();

    IRegistrationInfo* pInfo = NULL;
    hr = pTask->get_RegistrationInfo(&pInfo);
    if (SUCCEEDED(hr))
    {
        pInfo->put_Author(_bstr_t(L"Author"));  //计划任务Author，可改为Micsoft
        pInfo->Release();
    }

    // 设置触发器
    ITriggerCollection* pTriggers = NULL;
    hr = pTask->get_Triggers(&pTriggers);
    if (SUCCEEDED(hr))
    {
        ITrigger* pTrigger = NULL;
        hr = pTriggers->Create(TASK_TRIGGER_LOGON, &pTrigger); //TASK_TRIGGER_LOGON：用户登陆时触发

        

        pTriggers->Release();
    }

    // 设置操作（例如，运行 Notepad）
    IActionCollection* pActions = NULL;
    hr = pTask->get_Actions(&pActions);
    if (SUCCEEDED(hr))
    {
        IAction* pAction = NULL;
        hr = pActions->Create(TASK_ACTION_EXEC, &pAction);
        if (SUCCEEDED(hr))
        {
            std::cout << "success " << hr << std::endl;
            IExecAction* pExecAction = NULL;
            pAction->QueryInterface(IID_IExecAction, (void**)&pExecAction);

            pExecAction->put_Path(_bstr_t(L"C:\\test.exe")); //更改路径
            pExecAction->Release();
            pAction->Release();
        }
        pActions->Release();
    }

    // 注册任务
// 注册任务
    IRegisteredTask* pRegisteredTask = NULL;
    hr = pRootFolder->RegisterTaskDefinition(
        _bstr_t(L"Test Task"), //计划任务名称
        pTask,
        TASK_CREATE_OR_UPDATE,
        _variant_t(L"system"),
        _variant_t(),
        TASK_LOGON_GROUP,
        _variant_t(L""),
        &pRegisteredTask);

    if (FAILED(hr))
    {
        printf("Failed to register the task definition. Error code = %x\n", hr);
    }
    else
    {
        printf("Task successfully registered.\n");
        pRegisteredTask->Release();
    }

    // 释放资源
    pRootFolder->Release();
    pTask->Release();
    pService->Release();
    CoUninitialize();

    return 0;
}
