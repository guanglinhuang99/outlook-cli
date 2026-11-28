import Foundation
import NaturalLanguage

struct Request: Decodable {
    let language: String
    let texts: [String]
}

struct Result: Encodable {
    let revision: Int
    let dimension: Int
    let embeddings: [[Double]?]
}

do {
    let request = try JSONDecoder().decode(Request.self, from: FileHandle.standardInput.readDataToEndOfFile())
    guard let model = NLEmbedding.sentenceEmbedding(for: NLLanguage(rawValue: request.language)) else {
        exit(2)
    }
    let vectors = request.texts.map { model.vector(for: $0) }
    let result = Result(revision: model.revision, dimension: model.dimension, embeddings: vectors)
    FileHandle.standardOutput.write(try JSONEncoder().encode(result))
} catch {
    exit(4)
}
