import win32com.client

#### Create
# 創建 Task Scheduler 的 COM 物件
scheduler = win32com.client.Dispatch('Schedule.Service')

# 連接到本地的 Task Scheduler 服務
scheduler.Connect()

# 創建一個新的任務
root_folder = scheduler.GetFolder('\\')
task = scheduler.NewTask(0)

# 設置任務的基本屬性
task.RegistrationInfo.Description = '我的任務'
task.Settings.Enabled = True

# 設置任務的觸發器
trigger = task.Triggers.Create(win32com.client.constants.TaskTriggerTime)
trigger.StartBoundary = '2023-05-15T13:30:00'  # 任務開始運行的時間
trigger.Id = 'MyTrigger'
trigger.Enabled = True

# 設置任務的操作
action = task.Actions.Create(win32com.client.constants.TaskActionExecute)
action.Path = 'C:\\Python\\python.exe'
action.Arguments = 'C:\\Users\\User\\my_script.py'

# 將任務添加到 Task Scheduler 中
tasks_folder = root_folder.GetFolder('TaskName')
tasks_folder.RegisterTaskDefinition(
    'MyTask',  # 任務的名稱
    task,  # 要添加的任務
    win32com.client.constants.TASK_CREATE_OR_UPDATE,  # 如果任務已存在，則更新任務
    '',  # 執行任務的用戶名稱
    '',  # 執行任務的用戶密碼
    win32com.client.constants.TASK_LOGON_INTERACTIVE_TOKEN)  # 以交互式方式登錄用戶

print('Task created successfully.')

################
#### Cancelled

import win32com.client

# 創建 Task Scheduler 的 COM 物件
scheduler = win32com.client.Dispatch('Schedule.Service')

# 連接到本地的 Task Scheduler 服務
scheduler.Connect()

# 獲取指定名稱的任務
root_folder = scheduler.GetFolder('\\')
tasks_folder = root_folder.GetFolder('TaskName')
task = tasks_folder.GetTask('MyTask')

# 取消任務
task.Enabled = False
tasks_folder.DeleteTask('MyTask', 0)