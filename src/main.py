from constants import DEALERS, STOCK, logging
import pandas as pd
from pywhatkit import sendwhatmsg_instantly


def test():
    # Send a WhatsApp Message to a Contact at 1:30 PM
    pywhatkit.sendwhatmsg("+910123456789", "Hi", 13, 30)

    # Same as above but Closes the Tab in 2 Seconds after Sending the Message
    pywhatkit.sendwhatmsg("+910123456789", "Hi", 13, 30, 15, True, 2)

    # Send an Image to a Group with the Caption as Hello
    pywhatkit.sendwhats_image("AB123CDEFGHijklmn", "Images/Hello.png", "Hello")

    # Send an Image to a Contact with the no Caption
    pywhatkit.sendwhats_image("+910123456789", "Images/Hello.png")

    # Send a WhatsApp Message to a Group at 12:00 AM
    pywhatkit.sendwhatmsg_to_group("AB123CDEFGHijklmn", "Hey All!", 0, 0)

    # Send a WhatsApp Message to a Group instantly
    pywhatkit.sendwhatmsg_to_group_instantly("AB123CDEFGHijklmn", "Hey All!")

    # Play a Video on YouTube
    pywhatkit.playonyt("PyWhatKit")

    # Converting image to ASCII Art
    ascii_art = pywhatkit.image_to_ascii_art("image path")
    print(ascii_art)


def main():
    try:
        mobile = "unknown"
        df = pd.read_excel(STOCK, engine="openpyxl")
        message_columns = ["stock"]
        message = ""
        for _, row in df.iterrows():
            if row["stock"] != 0:
                row_msg = ", ".join(
                    str(row[col])
                    for col in df.columns
                    if col not in message_columns and pd.notnull(row[col])
                )
                message += row_msg + "\n" + "\n"
        message += "visit us at https://ecomsense.in \n"
        message += "reply with `STOP` to unsubscribe"
        df = pd.read_excel(DEALERS, index_col="mobile", engine="openpyxl")
        for mobile, row in df.iterrows():
            if pd.isna(row["stop"]):
                full_msg = "Dear " + row["salutation"] + ",\n" + "\n" + message
                sendwhatmsg_instantly(
                    phone_no="+91" + str(mobile),
                    message=full_msg,
                    tab_close=True,
                    close_time=3,
                )
                logging.info(f"message sent successfully to {mobile}")
    except Exception as e:
        print(f"{e} while sending whatsapp messages to {mobile}")


if __name__ == "__main__":
    main()
