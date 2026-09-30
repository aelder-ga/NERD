/* global Office, document */
const AUTO_OPEN_KEY = "Office.AutoShowTaskpaneWithDocument";

export async function configureAutoOpen() {
    const control = document.getElementById("nerd-auto-open");
    const container = document.getElementById("nerd-auto-open-control");
    const status = document.getElementById("nerd-auto-open-status");
    const settings = Office.context.document.settings;
    if (!control || !container || !status || !settings ||
        !Office.context.requirements.isSetSupported("AddInCommands", "1.1")) return;

    container.hidden = false;
    control.disabled = true;
    const saved = settings.get(AUTO_OPEN_KEY);
    control.checked = saved !== false;
    const persist = enabled => new Promise((resolve, reject) => {
        const previous = settings.get(AUTO_OPEN_KEY);
        settings.set(AUTO_OPEN_KEY, enabled);
        settings.saveAsync(result => {
            if (result.status === Office.AsyncResultStatus.Failed) {
                if (previous === undefined) settings.remove(AUTO_OPEN_KEY);
                else settings.set(AUTO_OPEN_KEY, previous);
                reject(new Error("Word could not save automatic opening. Check document editing permissions and try again."));
            } else resolve();
        });
    });
    try {
        if (saved !== true && saved !== false) await persist(true);
        status.textContent = "";
    } catch (error) {
        control.checked = false;
        status.textContent = error.message;
    } finally {
        control.disabled = false;
    }
    control.onchange = async () => {
        const enabled = control.checked;
        control.disabled = true;
        try {
            await persist(enabled);
            status.textContent = enabled ? "NERD will open with this document." : "Automatic opening is off for this document.";
        } catch (error) {
            control.checked = settings.get(AUTO_OPEN_KEY) === true;
            status.textContent = error.message;
        } finally {
            control.disabled = false;
        }
    };
}
