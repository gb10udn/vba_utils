# INFO: 240125 sys.stdout.write でなくても、print でも OK だった。
# import sys
# sys.stdout.write('stdout_output')


# INFO: 240125 end はデフォルトは \n。これが不要な場合は、空文字を指定する。
# print('print_output !!!', end='')


# INFO: 240125 繰り返し出力すると、\n で VBA には受け取られる。
for i in range(3):
    print(f'for i = {i}')


# import time
# from tqdm import tqdm
# for i in tqdm(range(10)):  # INFO: 240126 tqdm は標準出力に出てこなかった。。？
#     time.sleep(0.1)