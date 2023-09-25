import re
from typing import Callable

import openpyxl
import jieba
from wordcloud import WordCloud
from matplotlib import pyplot as plt


def get_statistics(msgs_path: str,
                   msg_filter: Callable[[str, str, str, str], bool] = None,
                   msg_preprocessor: Callable[[str], str] = None,
                   word_filter: Callable[[str], bool] = None,
                   speak_filter: Callable[[str, str], bool] = None) \
        -> tuple[list[str], dict[str, int], dict[str, int], dict[str, int], dict[str, str]]:
    """Read and analyse the given history message file of a QQ group, returns:

        - A list of split words.
        - A map from each date to message count in that day.
        - A map from each date when at least one anonymous message is present to anonymous message count in that day.
        - A map from each qq number to his/her speak count.
        - A map from each qq number to his/her latest card name.

        Arguments:

        - msgs_path -- path to the exported message history file of a QQ group in txt format.
        - msg_filter -- an optional predicate that accepts (message date, sender name, sender qq, message) and returns whether the message should be included.
        - msg_preprocessor -- an optional preprocessor that is applied to each message before it is processed.
        - word_filter -- an optional predicate that accepts a word and returns whether the word should be included in split word list.
        - speak_filter -- an optional predicate that accepts (sender name, sender qq) and returns whether the sender's speak count should be increased.
    """

    words = []
    total_counts = {}
    annoy_counts = {}
    speak_counts = {}
    card_names = {}

    with open(msgs_path, "r", encoding="utf-8") as f:
        prev_is_info_line = False
        prev_info_line = ""
        prev_date = ""
        for i in f:
            i = i.strip()
            match_result = re.search(r"^\d{4}-\d{2}-\d{2}", i)
            if match_result:
                prev_is_info_line = True
                prev_info_line = i
                prev_date = match_result.group(0)
            elif prev_is_info_line:
                prev_is_info_line = False

                name_qq = " ".join(prev_info_line.split(" ")[2:])
                left_parenthesis_index = name_qq.rfind("(")
                if left_parenthesis_index == -1:
                    left_parenthesis_index = name_qq.rfind("<")

                name = name_qq[:left_parenthesis_index]
                qq = name_qq[left_parenthesis_index + 1:-1]

                if msg_filter is None or msg_filter(prev_date, name, qq, i):
                    if prev_date in total_counts:
                        total_counts[prev_date] += 1
                    else:
                        total_counts[prev_date] = 1

                    if qq == "80000000":
                        if prev_date in annoy_counts:
                            annoy_counts[prev_date] += 1
                        else:
                            annoy_counts[prev_date] = 1

                    if speak_filter is None or speak_filter(name, qq):
                        if qq in speak_counts:
                            speak_counts[qq] += 1
                        else:
                            speak_counts[qq] = 1

                        card_names[qq] = name

                    if msg_preprocessor is not None:
                        i = msg_preprocessor(i)

                    words += [word for word in jieba.lcut(i)
                              if word_filter is None or word_filter(word)]
            else:
                prev_is_info_line = False

    return words, total_counts, annoy_counts, speak_counts, card_names


def save_counts_xlsx(total_counts: dict[str, int],
                     annoy_counts: dict[str, int],
                     output_path: str):
    wb = openpyxl.Workbook()
    ws = wb.worksheets[0]

    for i in total_counts:
        annoy_count = 0
        annoy_percent = 0

        if i in annoy_counts:
            annoy_count = annoy_counts[i]

        if total_counts[i] != 0:
            annoy_percent = 100 * annoy_count / total_counts[i]
        ws.append([i, total_counts[i], annoy_count, annoy_percent])

    wb.save(output_path)


def save_word_cloud_png(words: list[str],
                        output_path: str = None):
    while True:
        wc = WordCloud(width=800,
                       height=600,
                       prefer_horizontal=1,
                       font_path="./msyh.ttc",
                       colormap="tab10",
                       mode="RGBA",
                       background_color="#ffffff") \
            .generate(" ".join(words))

        # plt.rcParams["font.sans-serif"] = ["SimHei"]
        plt.imshow(wc, interpolation='bilinear')
        plt.axis("off")
        plt.show()

        wc.to_file(output_path)

        if input("Satisfied?") == "1":
            break


def save_speak_counts_xlsx(speak_counts: dict[str, int],
                           card_names: dict[str, str],
                           output_path: str):
    name_counts = [[card_names[qq], speak_counts[qq]] for qq in speak_counts]
    name_counts.sort(key=lambda x: x[1], reverse=True)

    wb = openpyxl.Workbook()
    for i in name_counts:
        wb.worksheets[0].append(i)

    wb.save(output_path)


def print_top_words(words: list[str], top_count: int):
    counts = {}
    for word in words:
        if word in counts:
            counts[word] += 1
        else:
            counts[word] = 1

    sorted = [[i, counts[i]] for i in counts]
    sorted.sort(key=lambda x: x[1], reverse=True)
    top_count = min(top_count, len(sorted))

    for item in sorted[:top_count]:
        print(item[0])


stop_words = set()
with open("./stopwords.dat", "r", encoding="utf-8") as f:
    for i in f:
        stop_words.add(i.strip())

jieba.add_word("出了")
jieba.add_word("工程硕博")
jieba.add_word("计算学部")


def qq_msg_filter_common(qq: str, msg: str) -> bool:
    return not (qq == "1000000" or
                "一条匿名消息被撤回" in msg or
                "撤回了一条成员消息" in msg or
                "撤回了一条消息" in msg or
                "被设为了精华消息" in msg or
                "不支持的消息类型" in msg or
                "请使用最新版手机QQ" in msg)


def qq_msg_filter_10days(date: str, name: str, qq: str, msg: str) -> bool:
    day = int(date.split("-")[2])
    return (date.startswith("2023-09-") and 14 <= day <= 23) and \
           qq_msg_filter_common(qq, msg)


def qq_msg_filter_sept_x(day: int):
    return lambda date, name, qq, msg: \
        date == "2023-09-" + "{:02d}".format(day) and \
        qq_msg_filter_common(qq, msg)


def qq_msg_preprocessor(msg):
    # remove mentions
    return re.sub("@.+ ", "", msg)


def qq_word_filter(word):
    return len(word) > 1 and word not in stop_words


def no_annoy_sys_speak_filter(name, qq):
    return qq != "80000000" and qq != "1000000"


words_10days, total_counts_10days, annoy_counts_10days, spk_counts, card_names = get_statistics(
    "./6.0-0924.txt",
    msg_filter=qq_msg_filter_10days,
    msg_preprocessor=qq_msg_preprocessor,
    word_filter=qq_word_filter,
    speak_filter=no_annoy_sys_speak_filter)

save_speak_counts_xlsx(spk_counts, card_names, "./6.0-2/speak_counts.xlsx")
save_counts_xlsx(total_counts_10days, annoy_counts_10days, "./6.0-2/10days.xlsx")
save_word_cloud_png(words_10days, "./6.0-2/10days.png")

# for i in range(14, 24):
#     words_day, total_counts_day, annoy_counts_day, _, _ = get_statistics(
#         "./6.0-0924.txt",
#         msg_filter=qq_msg_filter_sept_x(i),
#         msg_preprocessor=qq_msg_preprocessor,
#         word_filter=qq_word_filter)
#
#     print("2023-09-" + "{:02d}".format(i))
#     print_top_words(words_day, 9)
