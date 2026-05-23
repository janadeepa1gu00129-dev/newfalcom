// Mask phone numbers: keep first 3 and last 2 digits
function maskPhoneNumber(phone) {
  if (!phone) return "";
  // Ensure it's digits only before masking
  const digits = phone.toString().replace(/\D/g, "");
  if (digits.length < 10) return phone; // fallback if not a standard number
  return digits.replace(/(\d{3})\d{5}(\d{2})/, "$1*****$2");
}

async function loadBookings() {
  try {
    const dropboxLink = "https://dl.dropboxusercontent.com/scl/fi/zsr8s25h7khrqiq3hegtx/Booking.xlsx?rlkey=x8r0yq1n4a61w148hz97o4tl3&st=4nlcdxvv";

    const response = await fetch(dropboxLink);
    const arrayBuffer = await response.arrayBuffer();
    const workbook = XLSX.read(arrayBuffer, { type: "array" });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const data = XLSX.utils.sheet_to_json(sheet);

    const bookingList = document.getElementById("bookingList");
    bookingList.innerHTML = "";

    data.forEach((row, index) => {
      const statusClass = row.Status.toLowerCase() === "confirmed" ? "confirmed" : "pending";

      // Apply masking here
      const maskedPhone = maskPhoneNumber(row["Phone No."]);

      const card = document.createElement("div");
      card.className = "booking-card fade-in-section";
      card.style.transitionDelay = `${index * 0.2}s`; // stagger effect
      card.innerHTML = `
        <div class="booking-left">
          <div class="booking-icon">
            <img src="images/profile-icon.png" alt="Profile Icon" />
          </div>
          <div class="booking-details">
            <h3>${row.Name}</h3>
            <p>${maskedPhone}</p>
          </div>
        </div>
        <div class="booking-right">
          <div class="booking-info">
            <div>
              <strong>Facility</strong>
              <p>${row.Type}</p>
            </div>
            <div>
              <strong>Time</strong>
              <p>${row.Time}</p>
            </div>
          </div>
          <span class="status ${statusClass}">${row.Status}</span>
        </div>
      `;
      bookingList.appendChild(card);
    });

    document.getElementById("lastUpdated").textContent =
      "Last updated: " + new Date().toLocaleTimeString();

    // Trigger animation immediately after load
    setTimeout(() => {
      document.querySelectorAll(".fade-in-section").forEach(el => el.classList.add("visible"));
    }, 300);
  } catch (error) {
    console.error("Error loading bookings:", error);
  }
}

// Initial load + auto-refresh every 10 seconds
loadBookings();
setInterval(loadBookings, 10000);

// Scroll animation trigger
window.addEventListener("scroll", () => {
  const bookingsSection = document.getElementById("bookings");
  const rect = bookingsSection.getBoundingClientRect();
  if (rect.top < window.innerHeight - 100) {
    document.querySelectorAll(".fade-in-section").forEach(el => el.classList.add("visible"));
  }
});

// Button click smooth scroll + animation
document.querySelector(".btn.primary").addEventListener("click", (e) => {
  e.preventDefault();
  document.getElementById("bookings").scrollIntoView({ behavior: "smooth" });
  setTimeout(() => {
    document.querySelectorAll(".fade-in-section").forEach(el => el.classList.add("visible"));
  }, 400);
});
window.addEventListener("scroll", () => {
  document.querySelectorAll(".facility-card").forEach((card) => {
    const rect = card.getBoundingClientRect();
    if (rect.top < window.innerHeight - 100) {
      card.classList.add("visible");
    }
  });
});

window.addEventListener("load", () => {
  setTimeout(() => {
    document.querySelectorAll(".facility-card").forEach((card, index) => {
      card.style.transitionDelay = `${index * 0.2}s`;
      card.classList.add("visible");
    });
  }, 300);
});


