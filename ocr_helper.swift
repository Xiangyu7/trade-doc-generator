import Foundation
import Vision
import AppKit

// 用法: ocr_helper <图片1> <图片2> ...
// 调用 macOS 自带 Vision OCR，逐张识别，页间用换页符(\f)分隔，文字输出到 stdout。
guard CommandLine.arguments.count > 1 else {
    FileHandle.standardError.write("usage: ocr_helper <image>...\n".data(using: .utf8)!)
    exit(1)
}

for path in CommandLine.arguments.dropFirst() {
    guard let img = NSImage(contentsOfFile: path),
          let cg = img.cgImage(forProposedRect: nil, context: nil, hints: nil) else {
        print("\u{0C}")
        continue
    }
    let request = VNRecognizeTextRequest()
    request.recognitionLevel = .accurate
    request.usesLanguageCorrection = true
    request.recognitionLanguages = ["en-US", "zh-Hans"]
    let handler = VNImageRequestHandler(cgImage: cg, options: [:])
    do {
        try handler.perform([request])
        for obs in request.results ?? [] {
            if let top = obs.topCandidates(1).first {
                print(top.string)
            }
        }
    } catch {
        // 跳过出错的页
    }
    print("\u{0C}")
}
