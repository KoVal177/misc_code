from random import randrange

def display_board(board):
    line = '+-------+-------+-------+'
    cell_empty = '|       |       |       |'
    for i in range(3):
        cell_data = f'|   {board[i][0]}   |   {board[i][1]}   |   {board[i][2]}   |'
        print(line)
        print(cell_empty)
        print(cell_data)
        print(cell_empty)
    print(line)


def enter_move(board):
    input_success = False
    while not input_success:
        try:
            move = int(input('Enter your move: '))
            row = (move - 1) // 3
            col = (move - 1) % 3
            if not((0 <= row <= 2) and (0 <= col <= 2)):
                raise ValueError('') 
        except ValueError:
            print('Wrong input. Try again')
            continue
        if (row, col) not in free_fields:
            print('The field is taken. Try again')
            continue
        board[row][col] = 'O'
        input_success = True


def make_list_of_free_fields(board):
    free_fields.clear()
    for i in range(3):
        for j in range(3):
            if isinstance(board[i][j], int):
                free_fields.append((i, j))

def victory_for(board, sign):
    win_combo = sign * 3
    for i in range(3):
        combo1 = ''
        combo2 = ''
        for j in range(3):
            combo1 += str(board[i][j])
            combo2 += str(board[j][i])
        if win_combo == combo1 or win_combo == combo2:
            return True
    combo1 = str(board[0][0]) + str(board[1][1]) + str(board[2][2])
    combo2 = str(board[0][2]) + str(board[1][1]) + str(board[2][0])
    if win_combo == combo1 or win_combo == combo2:
        return True
    return False

def draw_move(board):
    if (1,1) in free_fields:
        board[1][1] = 'X'
    else:
        tmp = free_fields[randrange(len(free_fields))]
        board[tmp[0]][tmp[1]] = 'X'

free_fields = []
print(free_fields == [])
board = [[(i+j*3)+1 for i in range(3)] for j in range(3)]
make_list_of_free_fields(board)

while (free_fields != []):
    draw_move(board)
    make_list_of_free_fields(board)
    display_board(board)
    if victory_for(board, 'X'):
        print('Computer won!')
        break

    if (free_fields == []):
        print('It is a draw.')
        break

    enter_move(board)
    make_list_of_free_fields(board)
    display_board(board)
    if victory_for(board, 'O'):
        print('You won!')
        break
else:
    print('It is a draw.')

