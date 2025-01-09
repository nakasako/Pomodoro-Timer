import win32com.client
import time, datetime

speaker_cast_name = "弦巻マキ (日)"
work_time = 25 # 分
rest_time = 5 # 分
stretch_interval = 2 # 回に1回

class CeVIO:
    def __init__(self):
        self.cevio = win32com.client.Dispatch("CeVIO.Talk.RemoteService2.ServiceControl2")
        self.cevio.StartHost(False)
        self.talker = win32com.client.Dispatch("CeVIO.Talk.RemoteService2.Talker2V40")
        self.talker.Cast = speaker_cast_name

    def speak(self, text):
        print(text)
        state = self.talker.Speak(text)
        state.Wait()

    def close(self):
        self.cevio.StopHost()


def main():
    # CeVIOの起動
    talker = CeVIO()
    talker.speak("タイマー、起動完了です。")
    talker.speak("仕事を開始して下さい。")
    
    n = 0
    try:
        while True:
            # 仕事の開始
            time.sleep(work_time*60)

            # 現在時刻を取得
            now = datetime.datetime.now()

            # 時と分を取得
            now_hour = now.hour
            now_minute = now.minute

            # 時刻をアナウンス
            talker.speak(f"{now_hour}時{now_minute}分です")

            # 休憩またはストレッチ
            if n % stretch_interval == 0:
                talker.speak(f"{work_time}分経過しました。{rest_time}分間の休憩です。")
            else:
                talker.speak(f"{work_time}分経過しました。ストレッチの時間です。{rest_time}分間、体を伸ばしましょう。")

            # 休憩の開始
            time.sleep(rest_time*60)

            # 休憩の終了
            talker.speak(f"{rest_time}分経過しました。仕事に戻りましょう。")

            # カウントアップ
            n += 1

    except KeyboardInterrupt:
        talker.speak("タイマー、終了です。")
        talker.close()

if __name__ == '__main__':
    main()