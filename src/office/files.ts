export function getCurrentWorkbookUrl(): Promise<string> {
  return new Promise((resolve, reject) => {
    Office.context.document.getFilePropertiesAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value.url);
      } else {
        reject(new Error(result.error.message));
      }
    });
  });
}

export function requireWebUrl(value: string): URL {
  let url: URL;
  try {
    url = new URL(value);
  } catch {
    throw new Error("This favorite does not have a valid web URL.");
  }

  if (url.protocol !== "https:" && url.protocol !== "http:") {
    throw new Error("Only SharePoint, OneDrive, and other HTTP or HTTPS links can be opened.");
  }
  return url;
}

export function getContainingFolderUrl(fileUrl: string): string {
  const url = requireWebUrl(fileUrl);
  const lastSlash = url.pathname.lastIndexOf("/");
  if (lastSlash < 0) {
    throw new Error("A containing folder could not be determined for this favorite.");
  }
  url.pathname = url.pathname.substring(0, lastSlash + 1);
  url.search = "";
  url.hash = "";
  return url.toString();
}

export function openWebUrl(value: string): void {
  Office.context.ui.openBrowserWindow(requireWebUrl(value).toString());
}

export function getExcelOpenUrl(fileUrl: string): string {
  const url = requireWebUrl(fileUrl).toString().replace(/\|/g, "%7C");
  return `ms-excel:ofe|u|${url}`;
}

export function openWorkbookUrl(fileUrl: string): void {
  const link = document.createElement("a");
  link.href = getExcelOpenUrl(fileUrl);
  link.target = "_blank";
  link.rel = "noopener";
  link.hidden = true;
  document.body.append(link);
  link.click();
  link.remove();
}

export function openContainingFolderUrl(fileUrl: string): void {
  openWebUrl(getContainingFolderUrl(fileUrl));
}
