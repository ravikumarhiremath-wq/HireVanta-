
function continuePage() {
    const email = document.getElementById("email").value.trim();
    const message = document.getElementById("message");

    if (!email || !email.includes("@")) {
        message.textContent = "Please enter a valid email address.";
        return;
    }

    message.textContent =
        "Your page is working! Secure account verification will be added next.";
}
