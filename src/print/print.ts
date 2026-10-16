/** Printing (placeholder until the print engine lands). */
import type { App } from '../app';
export async function preparePrint(app: App): Promise<() => void> { void app; return () => undefined; }
export async function showPrintPreview(app: App): Promise<void> { void app; }
