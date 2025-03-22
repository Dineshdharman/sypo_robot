import pyttsx3

def speak(text):
    # Initialize the TTS engine
    engine = pyttsx3.init()

    # Set properties (optional)
    engine.setProperty('rate', 150)  # Speed of speech (words per minute)
    engine.setProperty('volume', 1.0)  # Volume level (0.0 to 1.0)

    # Get available voices
    voices = engine.getProperty('voices')

    # Set an Indian male voice (if available)
    for voice in voices:
        if "india" in voice.name.lower() or "indian" in voice.name.lower():  # Look for Indian voices
            if "male" in voice.name.lower():  # Look for a male voice
                engine.setProperty('voice', voice.id)
                break

    # Speak the text
    engine.say(text)
    engine.runAndWait()

# Example usage
speak("Welcome to the Computer Science Department Symposium.")