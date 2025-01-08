import win32com.client
import time, datetime

speaker_cast_name = "弦巻マキ (日)"
work_time = 25 # 分
rest_time = 5 # 分
stretch_interval = 2 # 回に1回

def main():
    # CeVIOの起動
    cevio = win32com.client.Dispatch("CeVIO.Talk.RemoteService2.ServiceControl2")
    cevio.StartHost(False)
    talker = win32com.client.Dispatch("CeVIO.Talk.RemoteService2.Talker2V40")
    talker.Cast = speaker_cast_name

    state = talker.Speak("タイマー、起動完了です。")
    state.Wait()

    state = talker.Speak("仕事を開始して下さい。")
    state.Wait()
    
    n = 0
    while True:
        # 仕事の開始
        time.sleep(work_time*60)

        # 現在時刻を取得
        now = datetime.datetime.now()

        # 時と分を取得
        now_hour = now.hour
        now_minute = now.minute

        # 時刻をアナウンス
        state = talker.Speak(f"{now_hour}時{now_minute}分です")
        state.Wait()

        # 休憩またはストレッチ
        if n % stretch_interval == 0:
            state = talker.Speak(f"{work_time}分経過しました。{rest_time}分間の休憩です。")
            state.Wait()
        else:
            state = talker.Speak(f"{work_time}分経過しました。ストレッチの時間です。{rest_time}分間、体を伸ばしましょう。")
            state.Wait()

        # 休憩の開始
        time.sleep(rest_time*60)

        # 休憩の終了
        state = talker.Speak(f"{rest_time}分経過しました。仕事に戻りましょう。")
        state.Wait()

        # カウントアップ
        n += 1

if __name__ == '__main__':
    main()
