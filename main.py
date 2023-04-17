import docx
from itertools import product
import random
import fractions as fr

# def style_header():
#     # изменение свойств шрифта и размера шрифта
#     font = run.font
#     font.name = 'Times New Roman'
#     font.size = docx.shared.Pt(16)
#     run.italic = True
#     # изменение выравнивания (по центру)
#     paragraph.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER

# def style_task():
#     font = run.font
#     font.name = 'Times New Roman'
#     font.size = docx.shared.Pt(16)
#     paragraph.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.JUSTIFY

tasks = {
    1: 'Игральная кость бросается три раза. Тогда вероятность того, что сумма выпавших очков не меньше шестнадцати, равна: ',
    2: 'Вероятность, что наудачу брошенная в круг точка окажется внутри вписанного в него квадрата равна',
    3: 'Сантехник обслуживает три дома. Вероятность того, что в течение часа потребуется его помощь в первом доме, равна 0,15; во втором – 0,25; в третьем – 0,2. Тогда вероятность  того, что в течение часа потребуется его помощь хотя бы в одном доме, рав-на:',
    4: 'Предприятие выплачивает 44 % всех зарплат разнорабочим, а 56 % – остальным. Вероятность того, что разнорабочий не получит зарплату в срок, равна 0,2; а для остальных эта вероят-ность составляет 0,1. Тогда вероятность того, что очередная зар-плата будет выдана в срок, равна:',
    5: 'Имеются четыре коробки, в которых сидят по 3 белых и по 7 черных котят, и шесть коробок, в которых сидят по 8 белых и по 2 черных котенка. Из наудачу взятой коробки вынимается один котенок, который оказался белым. Тогда вероятность того, что этого котенка достали из первой серии коробок, равна:',
    6: 'Дискретная случайная величина X задана законом рас-пределения вероятностей:',
    7: 'Дискретная случайная величина X задана законом рас-пределения вероятностей:'
}

answers = dict()


def t1():
    n, r = (6, 3)
    items = range(1, n + 1)
    arrangements = list(product(items, repeat=r))

    numbers = {10: 'десяти', 11: 'одиннадцати', 12: 'двенадцати', 13: 'тринадцати', 14: 'четырнадцати',
               15: 'пятнадцати', 16: 'шестнадцати'}
    choose = random.randrange(10, 16)
    task = tasks[1].replace('шестнадцати', numbers[choose], 1)
    # print(task)

    counter_target = 0
    counter_wrong1, counter_wrong2 = 0, 0
    for i in range(len(arrangements)):
        if sum(arrangements[i]) >= choose: counter_target += 1
        if choose != 16:
            if sum(arrangements[i]) >= choose + 1: counter_wrong1 += 1
            if sum(arrangements[i]) >= choose - 2: counter_wrong2 += 1
        else:
            if sum(arrangements[i]) >= choose - 1: counter_wrong1 += 1
            if sum(arrangements[i]) >= choose - 2: counter_wrong2 += 1

    p = fr.Fraction(counter_target, len(arrangements))
    w1 = fr.Fraction(counter_target - 2, len(arrangements))
    w2 = fr.Fraction(counter_wrong1, len(arrangements))
    w3 = fr.Fraction(counter_wrong2, len(arrangements))

    # print(P)

    return [task, p, w1, w2, w3]


if __name__ == '__main__':
    t1()
    document = docx.Document()

    # задание стиля для header
    style_header = document.styles.add_style('f_header', docx.enum.style.WD_STYLE_TYPE.CHARACTER)
    style_header.font.name = 'Times New Roman'
    style_header.font.size = docx.shared.Pt(16)
    style_header.font.italic = True

    # добавление параграфа с вариантом
    paragraph = document.add_paragraph()
    run = paragraph.add_run('Вариант 4')
    run.style = style_header
    run.font.bold = True
    paragraph.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER

    # добавление блока с фамилией и группой
    paragraph = document.add_paragraph()
    run = paragraph.add_run('\nФамилия ________________________ Группа __________')
    run.style = style_header
    paragraph.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER

    # задание стиля для заданий
    style_task = document.styles.add_style('f_tasks', docx.enum.style.WD_STYLE_TYPE.CHARACTER)
    style_task.font.name = 'Times New Roman'
    style_task.font.size = docx.shared.Pt(16)

    # блок заданий

    # задание 1
    paragraph = document.add_paragraph()
    run = paragraph.add_run('1. ')
    run.style = style_task
    run.bold = True

    task = t1()
    run = paragraph.add_run(task[0])
    run.style = style_task

    table = document.add_table(rows=1, cols=4)
    table.style = 'Table Grid'
    table.alignment = docx.enum.table.WD_TABLE_ALIGNMENT.CENTER

    table_font = table.style.font
    table_font.name = 'Times New Roman'
    table_font.size = docx.shared.Pt(16)

    for row in table.rows:
        for cell in row.cells:
            cell_font = cell.paragraphs[0].style.font
            cell_font.name = 'Times New Roman'
            cell_font.size = docx.shared.Pt(16)


    task_ans = task[1:]
    random.shuffle(task_ans)

    row_cells = table.rows[0].cells
    row_cells[0].text = f"а) {task_ans[0]};"
    row_cells[1].text = f"б) {task_ans[1]};"
    row_cells[2].text = f"в) {task_ans[2]};"
    row_cells[3].text = f"г) {task_ans[3]}."

    # for cell in row_cells:
    #     paragraphs = cell.paragraphs
    #     for paragraph in paragraphs:
    #         paragraph.style = document.styles['Table Content']
    #         run = paragraph.runs[0]
    #         run.font.name = 'Times New Roman'
    #         run.font.size = docx.shared.Pt(16)
    #         run.font.bold = False
    #         run.font.italic = False
    #         paragraph.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER

    # задание 2
    paragraph = document.add_paragraph()
    run = paragraph.add_run('2. ')
    run.style = style_task
    run.bold = True

    run = paragraph.add_run(tasks[2])
    run.style = style_task

    # задание 3
    paragraph = document.add_paragraph()
    run = paragraph.add_run('3. ')
    run.style = style_task
    run.bold = True

    run = paragraph.add_run(tasks[3])
    run.style = style_task

    # задание 4
    paragraph = document.add_paragraph()
    run = paragraph.add_run('4. ')
    run.style = style_task
    run.bold = True

    run = paragraph.add_run(tasks[4])
    run.style = style_task

    # задание 5
    paragraph = document.add_paragraph()
    run = paragraph.add_run('3. ')
    run.style = style_task
    run.bold = True

    run = paragraph.add_run(tasks[5])
    run.style = style_task

    # задание 6
    paragraph = document.add_paragraph()
    run = paragraph.add_run('6. ')
    run.style = style_task
    run.bold = True

    run = paragraph.add_run(tasks[6])
    run.style = style_task

    # задание 7
    paragraph = document.add_paragraph()
    run = paragraph.add_run('7. ')
    run.style = style_task
    run.bold = True

    run = paragraph.add_run(tasks[7])
    run.style = style_task

    document.save('text.docx')
