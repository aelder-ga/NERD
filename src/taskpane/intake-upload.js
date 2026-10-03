// Upload sessions fail on name conflicts: intake never replaces an existing file.
import { uploadUrl } from '../security/requests';
export async function uploadIntakeFile(file, { createSession, putChunk }) {
    if (!file?.name || /["*:<>?\\/|]/.test(file.name) || /[. ]$/.test(file.name)) throw new Error("The filename contains unsupported characters.");
    if (!Number.isSafeInteger(file.size) || file.size < 1 || file.size > 250 * 1024 * 1024) throw new Error("Choose a file between 1 byte and 250 MB.");
    const session = await createSession(file.name, { item: { "@microsoft.graph.conflictBehavior": "fail", name: file.name } });
    const target = uploadUrl(session?.uploadUrl);
    const chunkSize = 16 * 320 * 1024;
    for (let start = 0; start < file.size; start += chunkSize) {
        const end = Math.min(start + chunkSize, file.size);
        const response = await putChunk(target, file.slice(start, end), {
            "Content-Type": "application/octet-stream",
            "Content-Range": `bytes ${start}-${end - 1}/${file.size}`,
        });
        const body = await response.json().catch(() => ({}));
        if (!response.ok) throw new Error(body.error?.message || `Upload failed (${response.status}). Refresh files before retrying.`);
        if (end === file.size) {
            if (![200, 201].includes(response.status) || !body.id) throw new Error("Upload completion could not be verified. Refresh files before retrying.");
            return body;
        }
        if (response.status !== 202) throw new Error("Microsoft returned an unexpected upload response.");
    }
}
