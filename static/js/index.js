async function submitHandle(event) {
  event.preventDefault();

  const loader = document.getElementById("loader-container");
  const downloader = document.getElementById("downloader");
  const btn = document.getElementById("btn-submit");
  const mainForm = document.getElementById("main-form");
  const formData = new FormData(mainForm);

  downloader.style.display = "none";
  btn.disabled = true;
  loader.style.display = "flex";

  try {
    const response = await fetch("/generate", {
      method: "POST",
      body: formData,
    });

    if (response.ok) {
      const result = await response.json();

      downloader.href = `/download/${result.data}`;
      downloader.style.display = "flex";
    } else {
      alert("Failed to generate the file.");
    }
  } catch (error) {
    console.error("Error:", error);
    alert("An error occurred while submitting the form.");
  }
  loader.style.display = "none";
  btn.disabled = false;
}

document.getElementById("main-form").addEventListener("submit", submitHandle);

function changeFormat() {
  const formatSwitch = document.getElementById("switch");
  const formatLabel = document.getElementById("format-label");
  const formatInput = document.getElementById("format");
  if (formatSwitch.checked) {
    formatLabel.innerHTML = "PDF";
    formatInput.value = "pdf";
  } else {
    formatLabel.innerHTML = "WORD";
    formatInput.value = "word";
  }
}
