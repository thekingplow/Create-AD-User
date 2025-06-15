#v1.7 (20.06.2025)
#Developed by Danilovich M.D.
#

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.DirectoryServices


function Get-ADGroups {
    try {
        $ou = "OU=ExampleOU,DC=example,DC=com"  #Указать путь к вашей OU
        $groups = Get-ADGroup -Filter * -SearchBase $ou | Select-Object -ExpandProperty Name | Sort-Object
        return $groups
    } catch {
        Write-Host "Ошибка при получении групп из Active Directory: $_"
        return @()
    }
}


function Add-UserToGroups {
    param (
        [Parameter(Mandatory=$true)]
        [string]$userDN,
        [Parameter(Mandatory=$true)]
        [string[]]$groupNames
    )

    foreach ($groupName in $groupNames) {
        try {
            $group = Get-ADGroup -Identity $groupName
            Add-ADGroupMember -Identity $group -Members $userDN
        } catch {
            Write-Host "Ошибка при добавлении пользователя в группу {$groupName}: $_"
        }
    }
}


# Функция для получения групп из OU в Active Directory
function Get-ADGroupsITSBel {
    try {
        $ou = "OU=<Department>,OU=<Unit>,OU=<MainOU>,DC=<DomainName>,DC=local"
        $groups = Get-ADGroup -Filter * -SearchBase $ou | Select-Object -ExpandProperty Name | Sort-Object
        return $groups
    } catch {
        Write-Host "Ошибка при получении групп из организационной единицы ITS-Bel: $_"
        return @()
    }
}


function Add-UserToDepartmentGroup {
    param (
        [Parameter(Mandatory=$true)]
        [string]$userDN,
        [Parameter(Mandatory=$true)]
        [string]$departmentGroupName
    )

    try {
        # Путь поиска групп в OU
        $ou = "OU=<Department>,OU=<Unit>,OU=<MainOU>,DC=<DomainName>,DC=local"
        $group = Get-ADGroup -Filter { Name -eq $departmentGroupName } -SearchBase $ou

        if ($group) {
            Add-ADGroupMember -Identity $group -Members $userDN
            Write-Host "Пользователь добавлен в группу отдела: $departmentGroupName" -ForegroundColor Green
        } else {
            Write-Host "Группа отдела '$departmentGroupName' не найдена в OU 'ITS-Bel'" -ForegroundColor Red
        }
    } catch {
        Write-Host "Ошибка при добавлении пользователя в группу отдела '$departmentGroupName': $_" -ForegroundColor Red
    }
}


function Get-ADUsersFromOU {
    try {
        $ou = "OU=<Department>,OU=<Unit>,OU=<MainOU>,DC=<DomainName>,DC=local"
        $users = Get-ADUser -Filter * -SearchBase $ou -Property DisplayName | Select-Object -ExpandProperty DisplayName | Sort-Object
        return $users
    } catch {
        Write-Host "Ошибка при получении пользователей из Active Directory: $_"
        return @()
    }
}


function Generate-Password {
    $length = 12
    $chars = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789!@#$%^&*()_-+=<>?"
    $password = -join ((1..$length) | ForEach-Object { $chars[(Get-Random -Maximum $chars.Length)] })
    return $password
}


# Функция для получения всех OU из Active Directory
function Get-AllOU {
    $ous = Get-ADOrganizationalUnit -Filter * -Properties DistinguishedName | Select-Object -ExpandProperty DistinguishedName | Sort-Object
    return $ous
}















# Create the form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Add User to Active Directory"
$form.Size = New-Object Drawing.Size(1260, 910)
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle # Установка фиксированного размера формы
$form.MaximizeBox = $false # Отключение кнопки максимизации
$form.StartPosition = "CenterScreen"
$form.Font = New-Object System.Drawing.Font("Arial", 11, [System.Drawing.FontStyle]::Bold) # Устанавливаем стиль и размер шрифта для всех элементов формы

$scriptPath = $PSScriptRoot

# Установка иконки
$iconPath = Join-Path -Path $scriptPath -ChildPath "images\ad.ico" # Укажите путь к вашей иконке
$form.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($iconPath)

# Загружаем изображение из файла (замените путь на свой)
$imagePath = Join-Path -Path $scriptPath -ChildPath "images\bg.jpg"
$image = [System.Drawing.Image]::FromFile($imagePath)

# Устанавливаем изображение как фон формы
$form.BackgroundImage = $image
$form.BackgroundImageLayout = "Stretch"  # Растягиваем изображение на всю форму



# метка номера версии
$labelVersion = New-Object System.Windows.Forms.Label
$labelVersion.Text = "v1.7 (20.06.2025)"
$labelVersion.Location = New-Object System.Drawing.Point(0, 0)
$labelVersion.Font = New-Object System.Drawing.Font("Arial", 7.5, [System.Drawing.FontStyle]::Bold)  # Увеличение размера шрифта и жирный шрифт
$labelVersion.AutoSize = $true  # Автоматический размер под текст
$labelVersion.BackColor = [System.Drawing.Color]::Transparent  # Установка прозрачного фона
$labelVersion.ForeColor = [System.Drawing.Color]::White
$form.Controls.Add($labelVersion)



# Создание заголовка
$labelTitle = New-Object System.Windows.Forms.Label
$labelTitle.Text = "CREATE USER IN ACTIVE DIRECTORY"
$labelTitle.Location = New-Object System.Drawing.Point(300, 50)
$labelTitle.Font = New-Object System.Drawing.Font("Arial", 26, [System.Drawing.FontStyle]::Bold)  # Увеличение размера шрифта и жирный шрифт
$labelTitle.AutoSize = $true  # Автоматический размер под текст
$labelTitle.BackColor = [System.Drawing.Color]::Transparent  # Установка прозрачного фона
$labelTitle.ForeColor = [System.Drawing.Color]::White
$form.Controls.Add($labelTitle)



# Create FirstName label and textbox
$labelFirstName = New-Object System.Windows.Forms.Label
$labelFirstName.Text = "Имя"
$labelFirstName.Location = New-Object System.Drawing.Point(50, 152)
$labelFirstName.BackColor = [System.Drawing.Color]::Transparent  # Установка прозрачного фона
$labelFirstName.AutoSize = $true  # Автоматический размер под текст
$labelFirstName.ForeColor = [System.Drawing.Color]::White
$form.Controls.Add($labelFirstName)



# Создание метки для звёздочки
$labelStarFirstName = New-Object System.Windows.Forms.Label
$labelStarFirstName.Text = "*"
$labelStarFirstName.Top = 151
$labelStarFirstName.Left = ($labelFirstName.Left + $labelFirstName.Width + -3)
$labelStarFirstName.ForeColor = [System.Drawing.Color]::Red
$labelStarFirstName.BackColor = [System.Drawing.Color]::Transparent  # Установка прозрачного фона

# Задание ширины метки
$labelStarFirstName.AutoSize = $false  # Отключение автоматического размера
$labelStarFirstName.Width = 10  # Задание конкретной ширины
$form.Controls.Add($labelStarFirstName)


$textBoxFirstName = New-Object System.Windows.Forms.TextBox
$textBoxFirstName.Width = 150  # Задание конкретной ширины
$textBoxFirstName.Location = New-Object System.Drawing.Point(170, 150)


# Обрезка пробелов при выходе из поля
$textBoxFirstName.Add_Leave({
    $textBoxFirstName.Text = $textBoxFirstName.Text.Trim()
})

$form.Controls.Add($textBoxFirstName)




# Create LastName label and textbox
$labelLastName = New-Object System.Windows.Forms.Label
$labelLastName.Text = "Фамилия"
$labelLastName.Location = New-Object System.Drawing.Point(50, 202)
$labelLastName.BackColor = [System.Drawing.Color]::Transparent  # Установка прозрачного фона
$labelLastName.AutoSize = $true  # Автоматический размер под текст
$labelLastName.ForeColor = [System.Drawing.Color]::White
$form.Controls.Add($labelLastName)


# Создание метки для звёздочки
$labelStarLastName = New-Object System.Windows.Forms.Label
$labelStarLastName.Text = "*"
$labelStarLastName.Top = 201
$labelStarLastName.Left = ($labelLastName.Left + $labelLastName.Width + -3)
$labelStarLastName.ForeColor = [System.Drawing.Color]::Red
$labelStarLastName.BackColor = [System.Drawing.Color]::Transparent  # Установка прозрачного фона

# Задание ширины метки
$labelStarLastName.AutoSize = $false  # Отключение автоматического размера
$labelStarLastName.Width = 10  # Задание конкретной ширины
$form.Controls.Add($labelStarLastName)


$textBoxLastName = New-Object System.Windows.Forms.TextBox
$textBoxLastName.Width = 150  # Задание конкретной ширины
$textBoxLastName.Location = New-Object System.Drawing.Point(170, 200)


$textBoxLastName.Add_Leave({
    $textBoxLastName.Text = $textBoxLastName.Text.Trim()
})

$form.Controls.Add($textBoxLastName)


# Create Middle Name label and textbox
$labelMiddleName = New-Object System.Windows.Forms.Label
$labelMiddleName.Text = "Отчество"
$labelMiddleName.Location = New-Object System.Drawing.Point(50, 252)
$labelMiddleName.BackColor = [System.Drawing.Color]::Transparent  # Установка прозрачного фона
$labelMiddleName.AutoSize = $true  # Автоматический размер под текст
$labelMiddleName.ForeColor = [System.Drawing.Color]::White
$form.Controls.Add($labelMiddleName)


$textBoxMiddleName = New-Object System.Windows.Forms.TextBox
$textBoxMiddleName.Width = 150  # Задание конкретной ширины
$textBoxMiddleName.Location = New-Object System.Drawing.Point(170, 250)

$textBoxMiddleName.Add_Leave({
    $textBoxMiddleName.Text = $textBoxMiddleName.Text.Trim()
})

$form.Controls.Add($textBoxMiddleName)



# Create LoginName label and textbox
$labelLoginName = New-Object System.Windows.Forms.Label
$labelLoginName.Text = "Login Name"
$labelLoginName.Location = New-Object System.Drawing.Point(50, 302)
$labelLoginName.BackColor = [System.Drawing.Color]::Transparent  # Установка прозрачного фона
$labelLoginName.AutoSize = $true  # Автоматический размер под текст
$labelLoginName.ForeColor = [System.Drawing.Color]::White
$form.Controls.Add($labelLoginName)

$textBoxLoginName = New-Object System.Windows.Forms.TextBox
$textBoxLoginName.Width = 150  # Задание конкретной ширины
$textBoxLoginName.Location = New-Object System.Drawing.Point(170, 300)

$textBoxLoginName.Add_Leave({
    $textBoxLoginName.Text = $textBoxLoginName.Text.Trim()
})
$form.Controls.Add($textBoxLoginName)



# Создание метки для сообщения о проверке логина
$labelLoginCheck = New-Object System.Windows.Forms.Label
$labelLoginCheck.BackColor = [System.Drawing.Color]::Transparent  # Установка прозрачного фона
$labelLoginCheck.AutoSize = $true  # Автоматический размер под текст
#$labelLoginCheck.ForeColor = [System.Drawing.Color]::White
$labelLoginCheck.Location = New-Object System.Drawing.Point(180, 326)
$labelLoginCheck.Width = 150
$form.Controls.Add($labelLoginCheck)

# Обработчик события для проверки логина
$textBoxLoginName.Add_TextChanged({
    param ($sender, $e)
    
    $loginName = $textBoxLoginName.Text
    if (![string]::IsNullOrWhiteSpace($loginName)) {
        $existingUser = Get-ADUser -Filter { SamAccountName -eq $loginName }
        if ($existingUser) {
            $labelLoginCheck.Text = "Login существует"
            $labelLoginCheck.ForeColor = [System.Drawing.Color]::Red
        } else {
            $labelLoginCheck.Text = "Login свободен"
            $labelLoginCheck.ForeColor = [System.Drawing.Color]::LightGreen
        }
    } else {
        $labelLoginCheck.Text = ""
    }
})



# Создание метки для звёздочки
$labelStarLoginName = New-Object System.Windows.Forms.Label
$labelStarLoginName.Text = "*"
$labelStarLoginName.Top = 301
$labelStarLoginName.Left = ($labelLoginName.Left + $labelLoginName.Width + -3)
$labelStarLoginName.ForeColor = [System.Drawing.Color]::Red
$labelStarLoginName.BackColor = [System.Drawing.Color]::Transparent  # Установка прозрачного фона

# Задание ширины метки
$labelStarLoginName.AutoSize = $false  # Отключение автоматического размера
$labelStarLoginName.Width = 10  # Задание конкретной ширины
$form.Controls.Add($labelStarLoginName)



# Create Employee ID label and textbox
$labelEmployeeID = New-Object System.Windows.Forms.Label
$labelEmployeeID.Text = "Таб. № сотрудника"
$labelEmployeeID.BackColor = [System.Drawing.Color]::Transparent  # Установка прозрачного фона
$labelEmployeeID.AutoSize = $true  # Автоматический размер под текст
$labelEmployeeID.ForeColor = [System.Drawing.Color]::White
$labelEmployeeID.Location = New-Object System.Drawing.Point(50, 352)
$form.Controls.Add($labelEmployeeID)

$textBoxEmployeeID = New-Object System.Windows.Forms.TextBox
$textBoxEmployeeID.Width = 110  # Задание конкретной ширины
$textBoxEmployeeID.Location = New-Object System.Drawing.Point(210, 350)


$textBoxEmployeeID.Add_Leave({
    $textBoxEmployeeID.Text = $textBoxEmployeeID.Text.Trim()
})

$form.Controls.Add($textBoxEmployeeID)


# Добавление обработчика события KeyPress для текстового поля eployeeId
$textBoxEmployeeId.Add_KeyPress({
    param ($sender, $e)
    if ($e.KeyChar -notmatch '[0-9]') {
        $e.Handled = $true
    }
})



# Create Telephone Number label and textbox
$labelTelephoneNumber = New-Object System.Windows.Forms.Label
$labelTelephoneNumber.Text = "Номер телефона"
$labelTelephoneNumber.BackColor = [System.Drawing.Color]::Transparent  # Установка прозрачного фона
$labelTelephoneNumber.AutoSize = $true  # Автоматический размер под текст
$labelTelephoneNumber.ForeColor = [System.Drawing.Color]::White
$labelTelephoneNumber.Location = New-Object System.Drawing.Point(50, 402)
$form.Controls.Add($labelTelephoneNumber)

$textBoxTelephoneNumber = New-Object System.Windows.Forms.TextBox
$textBoxTelephoneNumber.Width = 130
$textBoxTelephoneNumber.Location = New-Object System.Drawing.Point(190, 400)
$textBoxTelephoneNumber.Text = "+375"  # Установка начального значения
$form.Controls.Add($textBoxTelephoneNumber)



# Добавление обработчика события Enter для текстового поля phoneNumber
$textBoxTelephoneNumber.Add_Enter({
    param ($sender, $e)
    if ($textBoxTelephoneNumber.Text -eq "") {
        $textBoxTelephoneNumber.Text = "+375"
    }
    $textBoxTelephoneNumber.SelectionStart = $textBoxTelephoneNumber.Text.Length
})

# Добавление обработчика события KeyPress для текстового поля phoneNumber
$textBoxTelephoneNumber.Add_KeyPress({
    param ($sender, $e)
    $char = $e.KeyChar.ToString()

    # Проверяем, является ли введенный символ цифрой или пробелом
    if ($char -match '[0-9\s]') {

        # Удаляем пробелы из текста
        $textWithoutSpaces = $textBoxTelephoneNumber.Text.Replace(" ", "")

        # Разрешаем ввод цифр только после маски +375
        if ($textWithoutSpaces.Length -ge 4) { # После +375 есть 4 символа
            if ($textWithoutSpaces.Length -eq 13) { # Максимальная длина +375 и 9 цифр
                $e.Handled = $true
            }
        }
       } else {
        # Запрещаем ввод других символов
        $e.Handled = $true
    }

    # Устанавливаем позицию курсора в конец текста
    $textBoxTelephoneNumber.SelectionStart = $textBoxTelephoneNumber.Text.Length
})



# Create Password label and textbox
$labelPassword = New-Object System.Windows.Forms.Label
$labelPassword.Text = "Пароль"
$labelPassword.Location = New-Object System.Drawing.Point(50, 452)
$labelPassword.BackColor = [System.Drawing.Color]::Transparent  # Установка прозрачного фона
$labelPassword.AutoSize = $true  # Автоматический размер под текст
$labelPassword.ForeColor = [System.Drawing.Color]::White
$form.Controls.Add($labelPassword)

# Создание метки для звёздочки
$labelStarPassword = New-Object System.Windows.Forms.Label
$labelStarPassword.Text = "*"
$labelStarPassword.Top = 451
$labelStarPassword.Left = ($labelPassword.Left + $labelPassword.Width + -3)
$labelStarPassword.ForeColor = [System.Drawing.Color]::Red
$labelStarPassword.BackColor = [System.Drawing.Color]::Transparent  # Установка прозрачного фона

# Задание ширины метки
$labelStarPassword.AutoSize = $false  # Отключение автоматического размера
$labelStarPassword.Width = 10  # Задание конкретной ширины
$form.Controls.Add($labelStarPassword)

$textBoxPassword = New-Object System.Windows.Forms.TextBox
$textBoxPassword.Location = New-Object System.Drawing.Point(190, 450)
$textBoxPassword.Width = 130
$textBoxPassword.PasswordChar = '*'
$form.Controls.Add($textBoxPassword)



# Создание кнопки для генерации пароля
$buttonGeneratePassword = New-Object System.Windows.Forms.Button
$buttonGeneratePassword.Text = "ГЕНЕРИРОВАТЬ"
$buttonGeneratePassword.Top = 485
$buttonGeneratePassword.Left = 200
$buttonGeneratePassword.Width = 110
$buttonGeneratePassword.Height = 26
$buttonGeneratePassword.BackColor = [System.Drawing.Color]::LightGray

# Установка курсора при наведении
$buttonGeneratePassword.Cursor = [System.Windows.Forms.Cursors]::Hand

$buttonGeneratePassword.Font = New-Object System.Drawing.Font("Arial", 7.8, [System.Drawing.FontStyle]::Bold)

# Установка стиля кнопки на Flat и настройка рамки
$buttonGeneratePassword.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$buttonGeneratePassword.FlatAppearance.BorderColor = [System.Drawing.Color]::DarkSeaGreen
$buttonGeneratePassword.FlatAppearance.BorderSize = 2

$form.Controls.Add($buttonGeneratePassword)

# Добавление обработчика события для кнопки генерации пароля
$buttonGeneratePassword.Add_Click({
    $generatedPassword = Generate-Password
    $textBoxPassword.Text = $generatedPassword
    #[System.Windows.Forms.MessageBox]::Show("Сгенерированный пароль: $generatedPassword", "Информация", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
})



# Путь к изображению глаза (одинаковое для открытого и закрытого состояния)
$imagePath = Join-Path -Path $scriptPath -ChildPath "images\eye.png"

$eyeImage = [System.Drawing.Image]::FromFile($imagePath)
$backgroundImage = $image.Clone()

# Создание кнопки для отображения/скрытия пароля
$buttonTogglePasswordVisibility = New-Object System.Windows.Forms.Button
$buttonTogglePasswordVisibility.Top = 450
$buttonTogglePasswordVisibility.Left = 140
$buttonTogglePasswordVisibility.Width = 32
$buttonTogglePasswordVisibility.Height = 25
$buttonTogglePasswordVisibility.Image = $eyeImage
$buttonTogglePasswordVisibility.ImageAlign = [System.Drawing.ContentAlignment]::MiddleCenter
#$buttonTogglePasswordVisibility.Padding = New-Object System.Windows.Forms.Padding(3, -10, 3, -2)  # Лево, верх, право, низ
$buttonTogglePasswordVisibility.Text = ""

# Установка курсора при наведении
$buttonTogglePasswordVisibility.Cursor = [System.Windows.Forms.Cursors]::Hand

$form.Controls.Add($buttonTogglePasswordVisibility)

# Добавление обработчика события для кнопки отображения/скрытия пароля
$buttonTogglePasswordVisibility.Add_Click({
 if ($textBoxPassword.PasswordChar -eq "*") {
        $textBoxPassword.PasswordChar = $null
    } else {
        $textBoxPassword.PasswordChar = "*"
    }
})



# Создание метки для флажка "Password never expires"
$labelPasswordNeverExpires = New-Object System.Windows.Forms.Label
$labelPasswordNeverExpires.Text = "Пароль никогда не истекает:"
$labelPasswordNeverExpires.Top = 530
$labelPasswordNeverExpires.Left = 50
$labelPasswordNeverExpires.AutoSize = $true
$labelPasswordNeverExpires.BackColor = [System.Drawing.Color]::Transparent  # Установка прозрачного фона
$labelPasswordNeverExpires.ForeColor = [System.Drawing.Color]::White
$form.Controls.Add($labelPasswordNeverExpires)

# Создание флажка "Password never expires"
$checkBoxPasswordNeverExpires = New-Object System.Windows.Forms.CheckBox
$checkBoxPasswordNeverExpires.Top = 532.5
$checkBoxPasswordNeverExpires.Left = 290
$checkBoxPasswordNeverExpires.BackColor = [System.Drawing.Color]::Transparent  # Установка прозрачного фона
$checkBoxPasswordNeverExpires.AutoSize = $true  # Автоматический размер под текст
$checkBoxPasswordNeverExpires.Checked = $true  # Установка флажка по умолчанию

$checkBoxPasswordNeverExpires.Cursor = [System.Windows.Forms.Cursors]::Hand

$form.Controls.Add($checkBoxPasswordNeverExpires)



# Создание метки и списка с выбором
$labelSelection = New-Object System.Windows.Forms.Label
$labelSelection.Text = "Группы:"
$labelSelection.Top = 150
$labelSelection.Left = 450
$labelSelection.BackColor = [System.Drawing.Color]::Transparent  # Установка прозрачного фона
$labelSelection.AutoSize = $true  # Автоматический размер под текст
$labelSelection.ForeColor = [System.Drawing.Color]::White
$form.Controls.Add($labelSelection)

$labelSelection.Font = New-Object System.Drawing.Font("Arial", 13.8, [System.Drawing.FontStyle]::Bold)


$listBoxSelection = New-Object System.Windows.Forms.ListBox
$listBoxSelection.Top = 200
$listBoxSelection.Left = 450
$listBoxSelection.Width = 250
$listBoxSelection.Height = 399
$listBoxSelection.SelectionMode = [System.Windows.Forms.SelectionMode]::MultiExtended

# Установка размера и стиля шрифта
$listBoxSelection.Font = New-Object System.Drawing.Font("Calibri", 11.2, [System.Drawing.FontStyle]::Regular)

$listBoxSelection.Cursor = [System.Windows.Forms.Cursors]::Hand

# Получение групп из Active Directory и добавление их в ListBox
$adGroups = Get-ADGroups
$listBoxSelection.Items.AddRange($adGroups)
$form.Controls.Add($listBoxSelection)



# Предопределенные группы для сотрудников офиса
$officeStaffGroups = @("Доступ к порталу", "Доступ к Wiki", "Пользователи в офисе", "Пользователи интернета", "Доступ к Wi-Fi", "Ресурсы Project Server", "Пользователи Project Server")

# Создание кнопки для выбора групп для сотрудников офиса
$buttonSelectOfficeGroups = New-Object System.Windows.Forms.Button
$buttonSelectOfficeGroups.Text = "Для сотрудников офиса"
$buttonSelectOfficeGroups.Top = 605
$buttonSelectOfficeGroups.Left = 450
$buttonSelectOfficeGroups.Width = 250
$buttonSelectOfficeGroups.Height = 40
$buttonSelectOfficeGroups.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
$buttonSelectOfficeGroups.ForeColor = [System.Drawing.Color]::Blue

# Установка курсора при наведении
$buttonSelectOfficeGroups.Cursor = [System.Windows.Forms.Cursors]::Hand

# Установка стиля кнопки на Flat и настройка рамки
$buttonSelectOfficeGroups.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$buttonSelectOfficeGroups.FlatAppearance.BorderColor = [System.Drawing.Color]::Silver
$buttonSelectOfficeGroups.FlatAppearance.BorderSize = 2

$form.Controls.Add($buttonSelectOfficeGroups)

# Добавление обработчика события Click для кнопки
$buttonSelectOfficeGroups.Add_Click({
    foreach ($group in $officeStaffGroups) {
        $index = $listBoxSelection.Items.IndexOf($group)
        if ($index -ge 0 -and -not $listBoxSelection.SelectedIndices.Contains($index)) {
            $listBoxSelection.SelectedIndices.Add($index)
        }
    }
})



# Предопределенные группы для сотрудников офиса
$branchStaffGroups = @("Доступ к Uploader", "Ограниченный доступ к порталу (филиалы)", "Доступ к Wiki")

# Создание кнопки для выбора групп для сотрудников офиса
$buttonSelectBranchGroups = New-Object System.Windows.Forms.Button
$buttonSelectBranchGroups.Text = "Для сотрудников филиалов и сервиса Минск"
$buttonSelectBranchGroups.Top = 655
$buttonSelectBranchGroups.Left = 450
$buttonSelectBranchGroups.Width = 250
$buttonSelectBranchGroups.Height = 44
$buttonSelectBranchGroups.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
$buttonSelectBranchGroups.ForeColor = [System.Drawing.Color]::Blue

# Установка курсора при наведении
$buttonSelectBranchGroups.Cursor = [System.Windows.Forms.Cursors]::Hand

# Установка стиля кнопки на Flat и настройка рамки
$buttonSelectBranchGroups.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$buttonSelectBranchGroups.FlatAppearance.BorderColor = [System.Drawing.Color]::Silver
$buttonSelectBranchGroups.FlatAppearance.BorderSize = 2

$form.Controls.Add($buttonSelectBranchGroups)

# Добавление обработчика события Click для кнопки
$buttonSelectBranchGroups.Add_Click({
    foreach ($group in $branchStaffGroups) {
        $index = $listBoxSelection.Items.IndexOf($group)
        if ($index -ge 0 -and -not $listBoxSelection.SelectedIndices.Contains($index)) {
            $listBoxSelection.SelectedIndices.Add($index)
        }
    }
})



# Create Job Title label and textbox
$labelJobTitle = New-Object System.Windows.Forms.Label
$labelJobTitle.Text = "Должность"
$labelJobTitle.Location = New-Object System.Drawing.Point(820, 152)
$labelJobTitle.BackColor = [System.Drawing.Color]::Transparent  # Установка прозрачного фона
$labelJobTitle.AutoSize = $true  # Автоматический размер под текст
$labelJobTitle.ForeColor = [System.Drawing.Color]::White
$form.Controls.Add($labelJobTitle)

# Создание метки для звёздочки
$labelStarJobTitle = New-Object System.Windows.Forms.Label
$labelStarJobTitle.Text = "*"
$labelStarJobTitle.Top = 151
$labelStarJobTitle.Left = ($labelJobTitle.Left + $labelJobTitle.Width + -3)
$labelStarJobTitle.ForeColor = [System.Drawing.Color]::Red
$labelStarJobTitle.BackColor = [System.Drawing.Color]::Transparent  # Установка прозрачного фона

# Задание ширины метки
$labelStarJobTitle.AutoSize = $false  # Отключение автоматического размера
$labelStarJobTitle.Width = 10  # Задание конкретной ширины
$form.Controls.Add($labelStarJobTitle)

$textBoxJobTitle = New-Object System.Windows.Forms.TextBox
$textBoxJobTitle.Location = New-Object System.Drawing.Point(960, 150)
$textBoxJobTitle.Width = 230


$textBoxJobTitle.Add_Leave({
    $textBoxJobTitle.Text = $textBoxJobTitle.Text.Trim()
})

$form.Controls.Add($textBoxJobTitle)



# Создание метки и текстового поля для Руководителя
$labelManagerName = New-Object System.Windows.Forms.Label
$labelManagerName.Text = "Руководитель:"
$labelManagerName.Top = 202
$labelManagerName.Left = 820
$labelManagerName.BackColor = [System.Drawing.Color]::Transparent  # Установка прозрачного фона
$labelManagerName.AutoSize = $true  # Автоматический размер под текст
$labelManagerName.ForeColor = [System.Drawing.Color]::White
$form.Controls.Add($labelManagerName)

# Создание метки для звёздочки
$labelStarManagerName = New-Object System.Windows.Forms.Label
$labelStarManagerName.Text = "*"
$labelStarManagerName.Top = 201
$labelStarManagerName.Left = ($labelManagerName.Left + $labelManagerName.Width + -3)
$labelStarManagerName.ForeColor = [System.Drawing.Color]::Red
$labelStarManagerName.BackColor = [System.Drawing.Color]::Transparent  # Установка прозрачного фона

# Задание ширины метки
$labelStarManagerName.AutoSize = $false  # Отключение автоматического размера
$labelStarManagerName.Width = 10  # Задание конкретной ширины
$form.Controls.Add($labelStarManagerName)


$textBoxManagerName = New-Object System.Windows.Forms.TextBox
$textBoxManagerName.Top = 200
$textBoxManagerName.Left = 960
$textBoxManagerName.Width = 230
$form.Controls.Add($textBoxManagerName)

# Добавление автодополнения и обработчика события TextChanged
$autoComplete = New-Object System.Windows.Forms.AutoCompleteStringCollection
$users = Get-ADUsersFromOU
$autoComplete.AddRange($users)

$textBoxManagerName.AutoCompleteCustomSource = $autoComplete
$textBoxManagerName.AutoCompleteMode = [System.Windows.Forms.AutoCompleteMode]::SuggestAppend
$textBoxManagerName.AutoCompleteSource = [System.Windows.Forms.AutoCompleteSource]::CustomSource

# Добавление обработчика события KeyPress для запрета ввода цифр и символов
$textBoxManagerName.Add_KeyPress({
    param($sender, $e)
    # Разрешаем только буквы (латиница и кириллица), пробел и дефис
    if (-not ($e.KeyChar -match '[а-яА-Яa-zA-Z \-]') -and -not [char]::IsControl($e.KeyChar)) {
        $e.Handled = $true
    }
})



# Создание метки и списка с выбором
$labelSelectionITSBel = New-Object System.Windows.Forms.Label
$labelSelectionITSBel.Text = "Отдел (ITS-Bel):"
$labelSelectionITSBel.Top = 250
$labelSelectionITSBel.Left = 820
$labelSelectionITSBel.BackColor = [System.Drawing.Color]::Transparent  # Установка прозрачного фона
$labelSelectionITSBel.AutoSize = $true  # Автоматический размер под текст
$labelSelectionITSBel.ForeColor = [System.Drawing.Color]::White
$form.Controls.Add($labelSelectionITSBel)

# Создание метки для звёздочки
$labelStarSelectionITSBel = New-Object System.Windows.Forms.Label
$labelStarSelectionITSBel.Text = "*"
$labelStarSelectionITSBel.Top = 250
$labelStarSelectionITSBel.Left = ($labelSelectionITSBel.Left + $labelSelectionITSBel.Width + -3)
$labelStarSelectionITSBel.ForeColor = [System.Drawing.Color]::Red
$labelStarSelectionITSBel.BackColor = [System.Drawing.Color]::Transparent  # Установка прозрачного фона

# Задание ширины метки
$labelStarSelectionITSBel.AutoSize = $false  # Отключение автоматического размера
$labelStarSelectionITSBel.Width = 10  # Задание конкретной ширины
$form.Controls.Add($labelStarSelectionITSBel)



$listBoxSelectionITSBel = New-Object System.Windows.Forms.ListBox
$listBoxSelectionITSBel.Top = 250
$listBoxSelectionITSBel.Left = 960
$listBoxSelectionITSBel.Width = 230
$listBoxSelectionITSBel.Height = 330
$listBoxSelectionITSBel.SelectionMode = [System.Windows.Forms.SelectionMode]::One

# Установка размера и стиля шрифта
$listBoxSelectionITSBel.Font = New-Object System.Drawing.Font("Calibri", 10.2, [System.Drawing.FontStyle]::Regular)

# Получение групп из ITS-Bel и добавление их в ListBox
$adGroupsITSBel = Get-ADGroupsITSBel
$listBoxSelectionITSBel.Items.AddRange($adGroupsITSBel)

$listBoxSelectionITSBel.Cursor = [System.Windows.Forms.Cursors]::Hand

$form.Controls.Add($listBoxSelectionITSBel)



# Создание метки и текстового поля для поиска OU
$labelSearchOU = New-Object System.Windows.Forms.Label
$labelSearchOU.Text = "Выбор OU:"
$labelSearchOU.Top = 602
$labelSearchOU.Left = 820
$labelSearchOU.AutoSize = $true
$labelSearchOU.ForeColor = [System.Drawing.Color]::White
$labelSearchOU.BackColor = [System.Drawing.Color]::Transparent  # Установка прозрачного фона
$form.Controls.Add($labelSearchOU)



$textBoxSearchOU = New-Object System.Windows.Forms.TextBox
$textBoxSearchOU.Top = 600
$textBoxSearchOU.Left = 960
$textBoxSearchOU.Width = 230
$form.Controls.Add($textBoxSearchOU)

# Установка начального значения в textBoxSearchOU
$initialValue = "OU=ExampleOU,DC=example,DC=com"
$textBoxSearchOU.Text = $initialValue

$listBoxOU = New-Object System.Windows.Forms.ListBox
$listBoxOU.Top = 650
$listBoxOU.Left = 960
$listBoxOU.Width = 230
$listBoxOU.Height = 100
$form.Controls.Add($listBoxOU)

# Получение всех OU и добавление их в ListBox
$allOUs = Get-AllOU
$listBoxOU.Items.AddRange($allOUs)

# Поиск индекса элемента, соответствующего начальному значению
$selectedIndex = $listBoxOU.FindStringExact($initialValue)
if ($selectedIndex -ne -1) {
    $listBoxOU.SelectedIndex = $selectedIndex
}

# Обработчик события TextChanged для TextBox поиска OU
$textBoxSearchOU.add_TextChanged({
    $searchText = $textBoxSearchOU.Text
    $ous = Get-AllOU | Where-Object { $_ -like "*$searchText*" }
    $listBoxOU.Items.Clear()
    $listBoxOU.Items.AddRange($ous)
})



# Создание элемента LinkLabel
$linkLabel = New-Object System.Windows.Forms.LinkLabel
$linkLabel.Text = "ПОДРОБНАЯ ИНСТРУКЦИЯ СОЗДАНИЯ УЧЕТНОЙ ЗАПИСИ"
$linkLabel.Top = 800
$linkLabel.Left = 50
$linkLabel.AutoSize = $true
$linkLabel.BackColor = [System.Drawing.Color]::Transparent  # Установка прозрачного фона
$form.Controls.Add($linkLabel)


# Убираем нижнее подчеркивание и устанавливаем цвет ссылки
$linkLabel.LinkBehavior = [System.Windows.Forms.LinkBehavior]::NeverUnderline
$linkLabel.LinkColor = [System.Drawing.Color]::LightGreen
$linkLabel.VisitedLinkColor = [System.Drawing.Color]::White
$linkLabel.ActiveLinkColor = [System.Drawing.Color]::White

# Установка ссылки
$link = "https://wiki.itsbel.by/pages/viewpage.action?pageId=169869792"
$linkLabel.Links.Add(0, $linkLabel.Text.Length, $link)

# Путь к Google Chrome (обновите путь, если у вас он другой)
$chromePath = "C:\Program Files\Google\Chrome\Application\chrome.exe"

# Добавление обработчика события для открытия ссылки в Google Chrome
$linkLabel.add_LinkClicked({
    param ($sender, $e)
    & $chromePath $e.Link.LinkData.ToString()
})

# Сохраняем цвет по умолчанию
$defaultColor = [System.Drawing.Color]::LightGreen
$hoverColor = [System.Drawing.Color]::Cyan

# Наведение — меняем цвет
$linkLabel.Add_MouseEnter({
    $linkLabel.LinkColor = $hoverColor
})

# Уход мыши — возвращаем цвет
$linkLabel.Add_MouseLeave({
    $linkLabel.LinkColor = $defaultColor
})




# Создание кнопки "Очистить"
$buttonClear = New-Object System.Windows.Forms.Button
$buttonClear.Text = "ОЧИСТИТЬ"
$buttonClear.Top = 780
$buttonClear.Left = 780
$buttonClear.Width = 130
$buttonClear.Height = 50      # Устанавливаем высоту кнопки
$buttonClear.BackColor = [System.Drawing.Color]::Silver
$buttonClear.ForeColor = [System.Drawing.Color]::Red
$buttonClear.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight  # Устанавливаем выравнивание текста по правому краю

# Установка курсора при наведении
$buttonClear.Cursor = [System.Windows.Forms.Cursors]::Hand

# Установка стиля кнопки на Flat и настройка рамки
$buttonClear.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$buttonClear.FlatAppearance.BorderColor = [System.Drawing.Color]::IndianRed
$buttonClear.FlatAppearance.BorderSize = 2

# Установка размера и стиля шрифта
$buttonClear.Font = New-Object System.Drawing.Font("Arial", 10.5, [System.Drawing.FontStyle]::Bold)

# Определение относительного пути к иконке
$brushPath = Join-Path -Path $scriptPath -ChildPath "images\brush.png"

# Загрузка и установка иконки для кнопки
$brush = [System.Drawing.Image]::FromFile($brushPath)
$buttonClear.Image = $brush

# Устанавливаем выравнивание иконки
$buttonClear.ImageAlign = [System.Drawing.ContentAlignment]::MiddleLeft

# Устанавливаем отступ справа для иконки
$buttonClear.Padding = New-Object System.Windows.Forms.Padding(5, 0, 5, 0)

# Добавление обработчика события для кнопки "Очистить"
$buttonClear.Add_Click({
    $form.Controls | ForEach-Object {
        if ($_ -is [System.Windows.Forms.TextBox]) {
            $_.Text = ""
        }
    }
})

$form.Controls.Add($buttonClear)




#Создание кнопки запуска AD Photo Edit
$buttonOpenADPhoto = New-Object System.Windows.Forms.Button
$buttonOpenADPhoto.Text = "AD Photo Edit"
$buttonOpenADPhoto.Top = 780
$buttonOpenADPhoto.Left = 995
$buttonOpenADPhoto.Width = 165
$buttonOpenADPhoto.Height = 50      # Устанавливаем высоту кнопки
$buttonOpenADPhoto.BackColor = [System.Drawing.Color]::Silver
$buttonOpenADPhoto.ForeColor = [System.Drawing.Color]::Sienna
$buttonOpenADPhoto.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight  # Устанавливаем выравнивание текста по правому краю

# Установка курсора при наведении
$buttonOpenADPhoto.Cursor = [System.Windows.Forms.Cursors]::Hand

# Установка стиля кнопки на Flat и настройка рамки
$buttonOpenADPhoto.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$buttonOpenADPhoto.FlatAppearance.BorderColor = [System.Drawing.Color]::Sienna
$buttonOpenADPhoto.FlatAppearance.BorderSize = 2

# Установка размера и стиля шрифта
$buttonOpenADPhoto.Font = New-Object System.Drawing.Font("Arial", 12.2, [System.Drawing.FontStyle]::Bold)

# Определение относительного пути к иконке
$photoPath = Join-Path -Path $scriptPath -ChildPath "images\add-photo.png"

# Загрузка и установка иконки для кнопки
$photo = [System.Drawing.Image]::FromFile($photoPath)
$buttonOpenADPhoto.Image = $photo

# Устанавливаем выравнивание иконки
$buttonOpenADPhoto.ImageAlign = [System.Drawing.ContentAlignment]::MiddleLeft

# Устанавливаем отступ справа для иконки
$buttonOpenADPhoto.Padding = New-Object System.Windows.Forms.Padding(5, 0, 5, 0)

# Обработчик нажатия кнопки
$buttonOpenADPhoto.Add_Click({
    $programPath = "C:\Program Files (x86)\Cjwdev\AD Photo Edit Free Edition\ADPhotoEdit.exe"
    
    if (Test-Path $programPath) {
        Start-Process $programPath
    } else {
        [System.Windows.Forms.MessageBox]::Show("Программа не найдена по пути:`n$programPath", "Ошибка", "OK", "Error")
    }
})

$form.Controls.Add($buttonOpenADPhoto)


# Create Add User button
$buttonAddUser = New-Object System.Windows.Forms.Button
$buttonAddUser.Text = "СОЗДАТЬ"
$buttonAddUser.Location = New-Object System.Drawing.Point(600, 780)
$buttonAddUser.Width = 130      # Устанавливаем ширину кнопки
$buttonAddUser.Height = 50      # Устанавливаем высоту кнопки
$buttonAddUser.BackColor = [System.Drawing.Color]::Silver  # Устанавливаем цвет фона кнопки
$buttonAddUser.ForeColor = [System.Drawing.Color]::Green      # Устанавливаем цвет текста кнопки
$buttonAddUser.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight  # Устанавливаем выравнивание текста по правому краю

# Установка курсора при наведении
$buttonAddUser.Cursor = [System.Windows.Forms.Cursors]::Hand

# Установка стиля кнопки на Flat и настройка рамки
$buttonAddUser.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$buttonAddUser.FlatAppearance.BorderColor = [System.Drawing.Color]::DarkGreen
$buttonAddUser.FlatAppearance.BorderSize = 2

# Определение относительного пути к иконке
$addPath = Join-Path -Path $scriptPath -ChildPath "images\add-user.png"

# Загрузка и установка иконки для кнопки
$add = [System.Drawing.Image]::FromFile($addPath)
$buttonAddUser.Image = $add

# Устанавливаем выравнивание иконки
$buttonAddUser.ImageAlign = [System.Drawing.ContentAlignment]::MiddleLeft

# Устанавливаем отступ справа для иконки
$buttonAddUser.Padding = New-Object System.Windows.Forms.Padding(5, 0, 10, 0)















# Создание пользователя при нажатии кнопки
# Обновление обработчика события для кнопки добавления пользователя
$buttonAddUser.Add_Click({
    $firstName = $textBoxFirstName.Text
    $lastName = $textBoxLastName.Text
    $loginName = $textBoxLoginName.Text
    $password = $textBoxPassword.Text
    $jobTitle = $textBoxJobTitle.Text
    $middleName = $textBoxMiddleName.Text
    $employeeID = $textBoxEmployeeID.Text
    $telephoneNumber = $textBoxTelephoneNumber.Text
    $selectedDepartment = $listBoxSelectionITSBel.SelectedItem
    $managerName = $textBoxManagerName.Text
    $selectedOU = $listBoxOUs.SelectedItem
    $company = "Name Company" #Указать свою компанию
    $fullName = "$lastName $firstName"
    $displayName = $fullName
    #$email = "$loginName@itsbel.by"
    $domain = "domain.com" #Указать свой домен
    $upnSuffix = "@$domain"
    $userPrincipalName = "$loginName$upnSuffix"
    $passwordNeverExpires = $checkBoxPasswordNeverExpires.Checked


    # Проверка на обязательные поля
    $missingFields = $false


    # Проверка имени
    if ([string]::IsNullOrWhiteSpace($firstName)) {
      $textBoxFirstName.BackColor = [System.Drawing.Color]::Red
      $missingFields = $true
    } else {
      $textBoxFirstName.BackColor = [System.Drawing.Color]::White
    }


    # Проверка фамилии
    if ([string]::IsNullOrWhiteSpace($lastName)) {
        $textBoxLastName.BackColor = [System.Drawing.Color]::Red
        $missingFields = $true
    } else {
        $textBoxLastName.BackColor = [System.Drawing.Color]::White
    }

    # Проверка логина
    if ([string]::IsNullOrWhiteSpace($loginName)) {
        $textBoxLoginName.BackColor = [System.Drawing.Color]::Red
        $missingFields = $true
    } else {
        $textBoxLoginName.BackColor = [System.Drawing.Color]::White
    }

    # Проверка Должности
    if ([string]::IsNullOrWhiteSpace($jobTitle)) {
        $textBoxJobTitle.BackColor = [System.Drawing.Color]::Red
        $missingFields = $true
    } else {
        $textBoxJobTitle.BackColor = [System.Drawing.Color]::White
    }

    # Проверка Руководителя
    if ([string]::IsNullOrWhiteSpace($managerName)) {
        $textBoxManagerName.BackColor = [System.Drawing.Color]::Red
        $missingFields = $true
    } else {
        $textBoxManagerName.BackColor = [System.Drawing.Color]::White
    }

    # Проверка пароля
    if ([string]::IsNullOrWhiteSpace($password)) {
        $textBoxPassword.BackColor = [System.Drawing.Color]::Red
        $missingFields = $true
    } else {
        $textBoxPassword.BackColor = [System.Drawing.Color]::White
    }

    # Проверка выбора отдела
    if ($listBoxSelectionITSBel.SelectedItem -eq $null) {
        $listBoxSelectionITSBel.BackColor = [System.Drawing.Color]::Red
        $missingFields = $true
    } else {
        $listBoxSelectionITSBel.BackColor = [System.Drawing.Color]::White
    }



    # Если есть пропущенные поля, показать сообщение об ошибке
    if ($missingFields) {
        [System.Windows.Forms.MessageBox]::Show("Поля Имя, Фамилия, Login Name, Должность, Руководитель, Отдел и Пароль обязательны для заполнения.", "Ошибка", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
     return
    }



    # Проверка пароля на соответствие требованиям
    if ($password.Length -lt 8 -or
        -not ($password -cmatch '[A-Z]') -or
        -not ($password -cmatch '[a-z]') -or
        -not ($password -cmatch '[\W_]')) {
            # Подсветить поле красным
        $textBoxPassword.BackColor = [System.Drawing.Color]::Red
        [System.Windows.Forms.MessageBox]::Show("Пароль должен быть не менее 8 символов, содержать как минимум одну заглавную букву, одну строчную букву и один специальный символ.", "Ошибка", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    } else {
        $textBoxPassword.BackColor = [System.Drawing.Color]::White
    }



    if ($listBoxOU.SelectedItem -ne $null) {
        $selectedOU = $listBoxOU.SelectedItem
        # Логика добавления пользователя в выбранное OU
        [System.Windows.Forms.MessageBox]::Show("Пользователь будет добавлен в OU: $selectedOU","Успешно",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information)
        # Здесь можно добавить код для добавления пользователя в Active Directory
    } else {
        [System.Windows.Forms.MessageBox]::Show("Пожалуйста, выберите OU.")
    }




    try {
        # Создание нового пользователя в выбранном OU
        $ou = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$selectedOU")
        $newUser = $ou.Children.Add("CN=$fullName", "user")
        $newUser.Properties["samAccountName"].Value = $loginName
        $newUser.Properties["userPrincipalName"].Value = $userPrincipalName
        $newUser.Properties["givenName"].Value = $firstName
        $newUser.Properties["sn"].Value = $lastName
        $newUser.Properties["displayName"].Value = $displayName
        $newUser.Properties["cn"].Value = $fullName
        $newUser.Properties["name"].Value = $fullName
        $newUser.Properties["title"].Value = $jobTitle
        $newUser.Properties["company"].Value = $company
        $newUser.Properties["department"].Value = $selectedDepartment

        if ($middleName) {
            $newUser.Properties["middleName"].Value = $middleName
        }

        if ($employeeID) {
            $newUser.Properties["employeeID"].Value = $employeeID
        }

        if ($telephoneNumber) {
            $newUser.Properties["telephoneNumber"].Value = $telephoneNumber
        }


        # Найти DistinguishedName руководителя
        $manager = Get-ADUser -Filter { DisplayName -eq $managerName } -Property DistinguishedName
        if ($manager) {
            $newUser.Properties["manager"].Value = $manager.DistinguishedName
        } else {
            throw "Руководитель не найден"
        }
        


        $newUser.CommitChanges()



        # Установка пароля
        $newUser.Invoke("SetPassword", $password)
        
       $userAccountControl = 512  # NORMAL_ACCOUNT
        if ($passwordNeverExpires) {
            $userAccountControl += 65536  # PASSWORD_NEVER_EXPIRES
        }
        $newUser.Properties["userAccountControl"].Value = $userAccountControl

        $newUser.CommitChanges()

        $selectedGroups = $listBoxSelection.SelectedItems
        if ($selectedGroups.Count -gt 0) {
            Add-UserToGroups -userDN $newUser.DistinguishedName -groupNames $selectedGroups
        }

         # Добавление пользователя в группу отдела
        if ($selectedDepartment) {
            Add-UserToDepartmentGroup -userDN $newUser.DistinguishedName -departmentGroupName $selectedDepartment
        }


        [System.Windows.Forms.MessageBox]::Show("Пользователь успешно добавлен!","Успешно",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information)
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Не удалось добавить пользователя: $_", "Ошибка", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }




    Write-Host "Имя: $firstName" -ForegroundColor Cyan
    Write-Host "Фамилия: $lastName" -ForegroundColor Cyan
    Write-Host "Отчество: $middleName" -ForegroundColor Cyan
    Write-Host "FullName: $fullname" -ForegroundColor Cyan
    Write-Host "Login Name: $loginName" -ForegroundColor Cyan
    Write-Host "Должность: $jobTitle" -ForegroundColor Cyan
    Write-Host "eployeeId: $employeeID" -ForegroundColor Cyan
    Write-Host "Телефон: $telephoneNumber" -ForegroundColor Cyan
    Write-Host "Компания: $company" -ForegroundColor Cyan
    Write-Host "Группы:" -ForegroundColor Cyan
        foreach ($group in $selectedGroups) {
            Write-Host $group -ForegroundColor Cyan
        }
    Write-Host "Отдел (ITS-Bel): $selectedDepartment" -ForegroundColor Cyan
    Write-Host "Руководитель: $managerName" -ForegroundColor Cyan
    Write-Host "Пароль: $password" -ForegroundColor Cyan
    Write-Host "OU: $selectedOU" -ForegroundColor Cyan
   # $form.Close()


})



$form.Controls.Add($buttonAddUser)


# Show the form
[void]$form.ShowDialog()
