Office.onReady(() => {
    document.getElementById("login-button").onclick = handleLogin;
});

async function handleLogin() {
    const user = document.getElementById('username').value;
    const pass = document.getElementById('password').value;

    const response = await fetch("https://macdemo.tesdadvts.org/apilogin", {
        method: "POST",
        headers: { 
            "Content-Type": "application/json" // <--- CRITICAL
        },
        body: JSON.stringify({ 
            username: user, 
            password: pass 
        })
    });

    if (response.ok) {
        const data = await response.json();
        localStorage.setItem("access_token", data.access_token);
        window.location.href = "dashboard.html";
    } else {
        const errorData = await response.json();
        alert("Login failed: " + (errorData.message || "Unknown error"));
    }
}