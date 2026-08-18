import streamlit as st


def check_login():

    # ----------------------------------------------------
    # Already logged in this session
    # ----------------------------------------------------

    if st.session_state.get(
        "authenticated",
        False
    ):
        return

    # ----------------------------------------------------
    # Load credentials from secrets.toml
    # ----------------------------------------------------

    try:

        correct_username = st.secrets[
            "auth"
        ]["username"]

        correct_password = st.secrets[
            "auth"
        ]["password"]

    except Exception:

        st.error(
            "Login credentials are not configured."
        )

        st.code(
            """
[auth]
username = "admin"
password = "your_password_here"
            """
        )

        st.stop()

    # ----------------------------------------------------
    # Login form
    # ----------------------------------------------------

    st.markdown(
        """
        <div style="text-align:center; margin-top:60px;">
            <div style="font-size:48px;">🕌</div>
            <h2 style="color:#164C3E; margin-bottom:0;">
                Dar Makkah International
            </h2>
            <p style="color:#6a706c;">
                Please log in to continue
            </p>
        </div>
        """,
        unsafe_allow_html=True
    )

    center_col = st.columns(
        [1, 1.2, 1]
    )[1]

    with center_col:

        with st.form("login_form"):

            username = st.text_input(
                "Username"
            )

            password = st.text_input(
                "Password",
                type="password"
            )

            submitted = st.form_submit_button(
                "🔓 Login",
                use_container_width=True,
            )

        if submitted:

            if (
                username == correct_username
                and password == correct_password
            ):

                st.session_state.authenticated = True
                st.rerun()

            else:

                st.error(
                    "Invalid username or password."
                )

    # ----------------------------------------------------
    # Stop the rest of the app from rendering
    # ----------------------------------------------------

    st.stop()


def logout_button():

   if st.button(
    "🚪 Logout",
    use_container_width=True,
    type="primary",
):

        st.session_state.authenticated = False
        st.rerun()
