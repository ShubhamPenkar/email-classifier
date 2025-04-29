import streamlit as st
import pickle
from sklearn.feature_extraction.text import CountVectorizer
import numpy as np
# Removed win32com.client and pythoncom
import threading

# Optional: Use a fallback speak_async that does nothing (cross-platform safe)
def speak_async(text):
    # Stub function that does nothing on non-Windows systems
    pass

# Load model and vectorizer
model = pickle.load(open('spam.pkl', 'rb'))
cv = pickle.load(open('vectorizer.pkl', 'rb'))

# Streamlit app
def main():
    st.title("📧 Email Spam Classification App")
    st.write("Built with Streamlit & Python")

    activities = ["Classification", "About"]
    choice = st.sidebar.selectbox("Select Activity", activities)

    if choice == "Classification":
        st.subheader("Enter your email content below:")
        msg = st.text_area("Your message:", height=150)

        if st.button("🔍 Classify"):
            if msg.strip() == "":
                st.warning("Please enter some text to classify.")
                return

            data = [msg]
            vec = cv.transform(data).toarray()
            result = model.predict(vec)

            if result[0] == 0:
                st.success("✅ This is Not a Spam Email.")
                speak_async("This is Not a Spam Email")
            else:
                st.error("🚫 This is a Spam Email!")
                speak_async("This is a Spam Email")

    elif choice == "About":
        st.subheader("About")
        st.markdown("""
        This app uses a machine learning model to detect whether a message is spam or not.
        - Built with **Scikit-learn**
        - Vectorized using **CountVectorizer**
        - Speech feedback via **SAPI (Windows only)** (disabled on this platform)
        """)

if __name__ == '__main__':
    main()
