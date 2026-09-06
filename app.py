import streamlit as st
from groq import Groq

# 1. Page Configuration
st.set_page_config(page_title="Groq AI Assistant", page_icon="⚡")
st.title("⚡ Groq AI Assistant")

# 2. Retrieve API key securely from Streamlit Secrets
groq_api_key = st.secrets.get("GROQ_API_KEY")

if not groq_api_key:
    st.error("GROQ_API_KEY is missing. Please configure secrets in Streamlit Cloud.")
    st.stop()

# Initialize the Groq Client
client = Groq(api_key=groq_api_key)

# 3. Maintain Session State for Chat History
if "messages" not in st.session_state:
    st.session_state.messages = []

# Display previous conversation history
for msg in st.session_state.messages:
    with st.chat_message(msg["role"]):
        st.markdown(msg["content"])

# 4. Handle User Input
if user_prompt := st.chat_input("Ask anything..."):
    # Render user prompt
    st.session_state.messages.append({"role": "user", "content": user_prompt})
    with st.chat_message("user"):
        st.markdown(user_prompt)

    # Request response from Groq API
    with st.chat_message("assistant"):
        message_placeholder = st.empty()
        full_response = ""

        try:
            # Stream response back to UI
            completion = client.chat.completions.create(
                model="llama-3.3-70b-versatile",
                messages=[
                    {"role": m["role"], "content": m["content"]}
                    for m in st.session_state.messages
                ],
                stream=True,
            )

            for chunk in completion:
                content = chunk.choices[0].delta.content or ""
                full_response += content
                message_placeholder.markdown(full_response + "▌")

            message_placeholder.markdown(full_response)

        except Exception as e:
            st.error(f"Error calling Groq API: {e}")

    # Store assistant response
    if full_response:
        st.session_state.messages.append({"role": "assistant", "content": full_response})