// QR code local (CoreImage, sem internet).  uso: qr <url> <saida.png> [px]
import CoreImage
import Foundation

let a = CommandLine.arguments
let size = a.count > 3 ? CGFloat(Double(a[3]) ?? 1200) : 1200
let f = CIFilter(name: "CIQRCodeGenerator")!
f.setValue(a[1].data(using: .utf8), forKey: "inputMessage")
f.setValue("Q", forKey: "inputCorrectionLevel")  // aguenta ~25% de dano/reflexo na TV
let qr = f.outputImage!.samplingNearest()
let scale = floor(size / qr.extent.width)
let img = qr.transformed(by: CGAffineTransform(scaleX: scale, y: scale))
try CIContext().writePNGRepresentation(of: img, to: URL(fileURLWithPath: a[2]), format: .RGBA8,
                                       colorSpace: CGColorSpace(name: CGColorSpace.sRGB)!)
print("\(a[2]): \(Int(img.extent.width))px, \(Int(qr.extent.width)) módulos")
