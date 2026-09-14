let statusTimeout: ReturnType<typeof setTimeout> | undefined;

export function setStatus(message: string): void {
  const status = document.getElementById("status");
  if (!status) return;
  if (statusTimeout) clearTimeout(statusTimeout);
  status.classList.remove("status-toast");
  status.textContent = message;
}

export function showStatusToast(message: string): void {
  const status = document.getElementById("status");
  if (!status) return;
  if (statusTimeout) clearTimeout(statusTimeout);
  status.textContent = message;
  status.classList.add("status-toast");
  statusTimeout = setTimeout(() => {
    status.textContent = "";
    status.classList.remove("status-toast");
    statusTimeout = undefined;
  }, 2400);
}
