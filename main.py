import docx

def style_header():
    # изменение свойств шрифта и размера шрифта
    font = run.font
    font.name = 'Times New Roman'
    font.size = docx.shared.Pt(16)
    run.italic = True
    # изменение выравнивания (по центру)
    paragraph.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER


if __name__ == '__main__':
    document = docx.Document()

    # добавление параграфа с вариантом
    paragraph = document.add_paragraph()
        # добавление текста в параграф
    run = paragraph.add_run('Вариант 4')
    run.bold = True
    style_header()

    # добавление блока с фамилией и группой
    paragraph = document.add_paragraph()
    run = paragraph.add_run('\nФамилия ________________________ Группа __________')
    style_header()

    tasks = [
        'Игральная кость бросается три раза. Тогда вероятность того, что сумма выпавших очков не меньше шестнадцати, равна: ',
        'Вероятность, что наудачу брошенная в круг точка окажет-ся внутри вписанного в него квадрата равна',
        '',
        '',

    ]



    document.save('text.docx')
