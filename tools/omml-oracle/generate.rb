# frozen_string_literal: true

require "digest"
require "fileutils"
require "json"
require "pathname"
require "plurimath"

abort "usage: bundle exec ruby generate.rb INPUT_DIRECTORY OUTPUT_DIRECTORY" unless ARGV.length == 2

input_root = Pathname.new(ARGV.fetch(0)).expand_path
output_root = Pathname.new(ARGV.fetch(1)).expand_path
abort "input directory does not exist: #{input_root}" unless input_root.directory?

results = []
["mathml", "latex", "unicodemath"].each do |format|
  FileUtils.rm_rf(output_root.join(format))
end

input_root.glob("**/*.omml").sort.each do |source|
  relative = source.relative_path_from(input_root)
  record = {
    source: relative.to_s.tr("\\", "/"),
    sha256: Digest::SHA256.file(source).hexdigest,
  }

  begin
    formula = Plurimath::Math.parse(source.read, :omml)
    outputs = {
      "mathml" => [".mathml", -> { formula.to_mathml }],
      "latex" => [".tex", -> { formula.to_latex }],
      "unicodemath" => [".txt", -> { formula.to_unicodemath }],
    }
    record[:outputs] = {}
    outputs.each do |format, (extension, render)|
      begin
        target = output_root.join(format, relative.sub_ext(extension))
        FileUtils.mkdir_p(target.dirname)
        target.write(render.call, mode: "w", encoding: "UTF-8")
        record[:outputs][format] = { status: "converted" }
      rescue StandardError => e
        record[:outputs][format] = {
          status: "error",
          error: "#{e.class}: #{e.message.lines.first&.strip}",
        }
      end
    end
    output_statuses = record[:outputs].values.map { |output| output[:status] }
    record[:status] = output_statuses.all? { |status| status == "converted" } ? "converted" : "partial"
  rescue StandardError => e
    record[:status] = "error"
    record[:error] = "#{e.class}: #{e.message.lines.first&.strip}"
  end
  results << record
end

FileUtils.mkdir_p(output_root)
output_root.join("manifest.json").write(
  JSON.pretty_generate(
    oracle: {
      plurimath_commit: "00c52783877b38f6b8e6e109f1803f96bb34fc62",
      omml_commit: "51d4abe5df58fe33a92df094971c5828c3459ffb",
    },
    fixtures: results,
  ) + "\n",
  mode: "w",
  encoding: "UTF-8",
)

puts "Processed #{results.length} fixtures " \
     "(#{results.count { |r| r[:status] == 'converted' }} complete, " \
     "#{results.count { |r| r[:status] == 'partial' }} partial, " \
     "#{results.count { |r| r[:status] == 'error' }} failed)."
