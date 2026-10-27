// Barcode scanning with ZXing: works in Chrome on Windows (webcam) and iPhone, which have
// no built-in BarcodeDetector. Frames are decoded in the page; no image is stored or sent.
import { BrowserMultiFormatReader, type IScannerControls } from '@zxing/browser';
import { cleanIsbn, isIsbn } from './isbn';

export async function startIsbnScan(
  video: HTMLVideoElement,
  onIsbn: (isbn: string) => void,
): Promise<IScannerControls> {
  const reader = new BrowserMultiFormatReader();
  return reader.decodeFromConstraints(
    { video: { facingMode: 'environment' } },
    video,
    (result, _error, controls) => {
      if (!result) return;
      const code = cleanIsbn(result.getText());
      if (isIsbn(code)) {
        controls.stop();
        onIsbn(code);
      }
    },
  );
}
