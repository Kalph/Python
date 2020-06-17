## 설명/설치

<br>

파이썬은 엑셀을 이용하여 DB에 값을 집어넣을 수 있는
<br>
여러 패키지를 제공한다.

<br>

먼저 모듈을 선택 전에 엑셀의 파일 확장자 형태를 살펴봐야 한다.
<br>

.xlrd로 끝나는 엑셀 파일을 다룰 수 있는 모듈인 pymssql은 .xls로 끝나는 엑셀 파일을 다룰 수 없다.
<br>
.xls는 xlrd와 같은 파이썬 패키지가 다룰 수 있다. (참고로 xlrd 패키지는 xlsx, xls 둘 다 다룰 수 있다.)
<br>
xlrd는 리스트로 값을 뽑아올 수 있다.

<br>

설치를 비롯한. 방법은 구글링을 통해서 해결한다.
<br>
pymssql 설치 시 오류가 발생할 수 있는데 여기를 참고한다.

<br>

 
[python pymssql 오류](https://blog.naver.com/PostView.nhn?blogId=kjskhj04366&logNo=221828675761&categoryNo=7&parentCategoryNo=0&viewDate=&currentPage=1&postListTopCurrentPage=&from=postList&userTopListOpen=true&userTopListCount=5&userTopListManageOpen=false&userTopListCurrentPage=1)

<br>

설치는 pip 명령어를 이용한다.(파이참 사용)

<br>

```py3
pip install pymssql
pip install xlrd
```
<br>

## 파일 읽어들이기

<br>

엑셀의 값을 뽑아올 때 엑셀이 한 폴더에 있는 경우가 아니라면 폴더 안의 엑셀을 일일히 하나의 폴더에 집어넣을 수는 없는 노릇이다.

<br>

외장함수 os를 이용한다.

<br>

```py3
import os

file_names = os.listdir(path)

for dir_path, dir_names, file_names in os.walk(path):
    if (goCount == 0):
        break
    file_path = "{0}".format(dir_path)
    file = "{0}".format(file_names)
    file_names = os.listdir(file_path)
    for file_name in file_names:
        full_filename = os.path.join(path, file_name)
        if file_name.find('xlsx') != -1 or file_name.find('xls') != -1:
            if file_name.find('~$') == -1 and file_name.find('iso') == -1 and file_name.find('zip') == -1 and file_name.find('-$') == -1:
                print('\n\n')
                print("파일 경로 = " + file_path)
                print("경로 내 폴더 = {0}".format(dir_names))
                print("file_name : " + file_name)
                # print(file_path + "/" + file_name)
```

<br>

위와 같은 코드의 형태는 최상위 폴더에서 모든 폴더로 접근 - 하나의 엑셀 파일씩 읽어들이며 반복하는 루프문이다.

<br>

필요하다면 임의로 수정해서 사용하면 된다.

<br>

## pymssql, xlrd 사용하기

<br>

### pymssql

```py3
import pymssql
conn = pymssql.connect(host='IP', user='username',
                       password='password', database='dbname')
cursor = conn.cursor()
cursor.execute() 
conn.commit()
```

<br>

### xlrd

<br>

```py3
import xlrd
wb = xlrd.open_workbook(file_path + "/" + file_name)
ws = wb.sheet_by_name('Sheet1')
```

<br>

* xlrd는 엑셀에서 값을 뽑아오는 것만 가능하다 결국 DB에 값을 집어넣을때는 pymssql을 사용한다. xlrd는 xls 엑셀 파일을 읽어들어야 하는 경우가 아니면 pymssql을 사용해야 한다.

<br>

## 기타 알고리즘

<br>

DB에 값을 집어넣을때 '' 의 형태보다 'null'로 집어넣는 경우가 더 보기 좋은것 같다.

그리고 DB에 들어가는 값이 aa'b 라면 '(작은따옴표)로 감싸서 집어넣을때 오류가 발생하게 된다.

<br>

이 경우를 방지 하기 위한 알고리즘이다.

<br>

```py3
def del_col(n):  #
    tem=""
    while n != '':
        if n.find("'") != -1:
            if len(n) == n.find("'")+1:
                tem += n[:n.find("'")+1]+"'"
                n= n[n.find("'")+1:]
                if n.find("'") == -1:
                    break
            elif n[n.find("'")+1] != "'":
                tem += n[:n.find("'") + 1] + "'"
                n= n[n.find("'") + 1:]
        if n[n.find("'")+1] == "'":
            n = n[:n.find("'")]+ '"' + n[n.find("'")+2:]
        if n.find("'") == -1:
            break
    return str(tem+n)


def nul_col(n):
    for i in range(0, len(n)):
        if n[i] == '' or n[i] == '-':
            n[i] = 'null'
    return n
```

<br>

del_col은 '가 들어간 문자열에 있을 경우 해당 문자열의'가 있는 위치 뒤에 '를 하나 더 붙여준다.

예를 들어 a'b -> a''b로 DB에 집어넣도록 설정한다. 

만약 "가 붙어있다면 '(작은따옴표)로 감싸게 될 경우 문제가 생기지 않으니 무시한다.

<br>

nul_col은 '' 이거나 '-'의 문자열을 null로 변경하여 리턴해준다.

<br>

엑셀의 값을 뽑아오려면 엑셀의 형태(셀 위치)까지도 완벽하게 동일해야 한다.

만일 이 점에서 문제가 생긴다면 값을 뽑아오는게 매우 힘들어지거나 코드가 기하 급수적으로 늘어난다.(일일히 처리해야 되서 ㅠㅠ)
