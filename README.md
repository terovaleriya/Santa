# Santa Post Tutorial

Все невероятно просто! 

- `matching.py`

    Работает с данными таблицами и присваивает каждому участнику подопечного. 

    Чтобы изменить рабочую таблицу, введите ее имя в переменную `data_table` (сейчас там указана `test.xlsx`). 
    
    Результат перемешивания будет записан в таблицу, указанную в `matching_table`.
    
- `emails.py`

    Вызывает `matching.py` при каждом запуске (то есть каждый раз обновляет `matching_table`).
    
    Собственно, отправляет участникам информацию об их подопечных.   

    ВНИМАНИЕ: на всякий случай, строка, в которой происходит отправка email, закомментирована. 
    
    На данный момент на почту участника отправляется одно из следующих сообщений:
    
        Привет, Платон!

        Твоего подопечного зовут Валерия Терова. Он учится на 2 курсе.
        Вот, что он оставил в качестве пожелания: "Спать".
        Его адрес: Дудко 24, 45678.
        К сожалению, твой подопечный не оставил ссылку на социальную сеть :(
        У нас есть номер телефона твоего подопечного, оставь его транспортной компании, так ему будет проще отслеживать полылку: +7-921-873-6059.

        Мы желаем тебе удачи и поздравляем с наступающим Новым годом!  
        
     ИЛИ
     
        Привет, Степан!
        
        Твоего подопечного зовут Платон Платонов. 
        Вот, что он оставил в качестве пожелания: "Есть".
        Его адрес: Седова 24, 456789.
        Ссылка на социальную сеть: https://vk.com/terovaleriya.
        Мы не располагаем номером его телефона.
        
        Мы желаем тебе удачи и поздравляем с наступающим Новым годом!