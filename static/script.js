let sessionId = 'default';

function updateOrdersDisplay(orders) {
    const ordersList = document.getElementById('ordersList');
    const orderItems = Object.entries(orders);
    
    if (orderItems.length === 0) {
        ordersList.innerHTML = '<div class="empty-orders">Nenhum pedido ainda</div>';
        return;
    }
    
    ordersList.innerHTML = orderItems.map(([product, quantity]) => `
        <div class="order-item">
            <span class="product-name">${product}</span>
            <span class="product-quantity">${quantity}</span>
        </div>
    `).join('');
}

function sendMessage() {
    const input = document.getElementById('messageInput');
    const message = input.value.trim();
    
    if (!message) return;
    
    // Add user message to chat
    addMessage(message, 'user');
    input.value = '';
    
    // Show typing indicator
    showTypingIndicator();
    
    // Send to server
    fetch('/send_message', {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json',
        },
        body: JSON.stringify({
            message: message,
            session_id: sessionId
        })
    })
    .then(response => response.json())
    .then(data => {
        hideTypingIndicator();
        addMessage(data.response, 'bot');
        updateOrdersDisplay(data.orders);
    })
    .catch(error => {
        hideTypingIndicator();
        addMessage('Erro de conexão. Tente novamente.', 'bot');
        console.error('Error:', error);
    });
}

function addMessage(text, sender) {
    const chatMessages = document.getElementById('chatMessages');
    const messageDiv = document.createElement('div');
    messageDiv.className = `message ${sender}-message`;
    messageDiv.textContent = text;
    chatMessages.appendChild(messageDiv);
    chatMessages.scrollTop = chatMessages.scrollHeight;
}

function showTypingIndicator() {
    document.getElementById('typingIndicator').style.display = 'block';
    document.getElementById('chatMessages').scrollTop = document.getElementById('chatMessages').scrollHeight;
}

function hideTypingIndicator() {
    document.getElementById('typingIndicator').style.display = 'none';
}

function downloadExcel() {
    window.open('/download_excel', '_blank');
}

// Enter key to send message
document.getElementById('messageInput').addEventListener('keypress', function(e) {
    if (e.key === 'Enter') {
        sendMessage();
    }
});

// Load initial orders when page loads
document.addEventListener('DOMContentLoaded', function() {
    fetch(`/get_orders?session_id=${sessionId}`)
        .then(response => response.json())
        .then(orders => updateOrdersDisplay(orders));
});