const keyInputs = document.getElementById("key-inputs");
const addKeyButton = document.getElementById("add-key");
let keyCounter = 1;

addKeyButton.addEventListener("click", () => {
  if (keyCounter >= 20) return;

  keyCounter++;

  const label = document.createElement("label");
  label.htmlFor = `key_${keyCounter}`;
  label.textContent = `Key ${keyCounter}:`;

  const input = document.createElement("input");
  input.type = "text";
  input.id = `key_${keyCounter}`;
  input.name = `key_${keyCounter}`;
  input.placeholder = `Key ${keyCounter}`;

  keyInputs.appendChild(label);
  keyInputs.appendChild(input);
});

let keyIndex = 0;

function addKey() {
  const keysContainer = document.getElementById('keys');
  const formGroup = document.createElement('div');
  formGroup.className = 'form-group';
  formGroup.id = `key-group-${keyIndex}`;

  const label = document.createElement('label');
  label.htmlFor = `key_${keyIndex}`;
  label.innerText = `Key ${keyIndex + 1}:`;

  const input = document.createElement('input');
  input.type = 'text';
  input.className = 'form-control';
  input.id = `key_${keyIndex}`;
  input.name = `key_${keyIndex}`;

  formGroup.appendChild(label);
  formGroup.appendChild(input);
  keysContainer.appendChild(formGroup);

  keyIndex++;
}

function removeKey() {
  if (keyIndex > 0) {
    keyIndex--;
    const keysContainer = document.getElementById('keys');
    const formGroup = document.getElementById(`key-group-${keyIndex}`);
    keysContainer.removeChild(formGroup);
  }
}

// Create an example key on page load
document.addEventListener('DOMContentLoaded', () => {
  addKey();
});

