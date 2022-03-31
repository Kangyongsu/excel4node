npm excel4node (기존에쓰던 xlsx는 셀변형하는게 불편한 점이 많아서 제가 한거 적어놓습니다 하시다가 괜찮다 싶은거 있으면 바로바로 추가부탁드립니다!!)

1.  모듈가져오기

    const xl = require("excel4node")

2.  객체생성

    const wb= new xl.Workbook()

3.  시트명생성

    const ws = wb.addWorksheet('sheet1')

4.  스타일생성

    const myStyle = wb.createStyle({

        font: {
            size : 10
            color: '00FF00',
        },
        alignment: {
            horizontal: 'center', (가운데정렬)
        },

    });

5.  row column 길이조절(데이터 길이가 길 셀처리하실때 사용하세욥)

    ws.column(3).setWidth(50)  
    =>3번째 column 가로길이를 50으로 설정

    ws.row(1).setHeight(20)  
    => 1번째 row 높이를 20으로 조정

6.  셀 커스터마이징

    (1)row column 지정
    ws.cell(2,10).string("설문일자").style({alignment:{horizontal:'center'}}).style({font: {size: 100}});  
    => 2번쨰행과 10번째열 위치에 문자열 "설문일자"라는 string타입을 넣음 + 가운데정렬 스타일 + 글자크기 100

    ws.cell(4,19).number(20).style({alignment:{horizontal:'center'}})  
    => 4번째행과 19번째열 위치에 20이라는 정수형타입을 넣음

    (2)셀 병합
    ws.cell(1,5,3,10,true).string("자격체크").style({alignment:{horizontal:'center',vertical:"center"}}).style({font: {size: 40}})  
    =>첫번째 파라미터 = 시작하는 행위치  
    =>2번째 파라미터 = 시작하는 열위치  
    =>3번쨰 파라미터 = 끝내는 행위치  
    =>4번째 파라미터 = 끝나는 열위치  
    =>5번째 파라미터 true로 설정안하면 적용이 안됩니다.

    ![예시1](./%EC%BA%A1%EC%B2%98.PNG)

7.  엑셀 다운로드

    wb.write('파일명',res)

ps 그 외에 정말 좋은 기능들 많은데 https://www.npmjs.com/package/excel4node 여기에서 참고하시고
다들 업데이트해서 귀찮은 엑셀작업 빠르게 끝냅시다

#프리것 화이팅!!
