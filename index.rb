class Vazao
    def initialize(usina, file_name)
        @usina = usina
        @chave_usina = nil
        @file_name = file_name
    end

    def read_file
        out_file = File.new("#{@usina}.txt", "w")

        sheet = read_sheet()
        sheet.to_a.shift(4)
        sheet.each_with_index do |row,index|
            get_chave_usina(row)
            puts_file(row, out_file)
        end

        out_file.close
    end

    def read_sheet
        require 'spreadsheet'    
        book = Spreadsheet.open(@file_name)
        book.worksheet 0
    end

    def puts_file(row, out_file)
        unless @chave_usina.nil?
            unless row[@chave_usina].nil? || row[0].nil?
                out_file.puts("#{row[0]};#{row[@chave_usina].to_i}")
            end
        end
    end
    
    def get_chave_usina(row)
        row.each_with_index do |cell,key|
            if cell.to_s.include? "(#{@usina})"
                @chave_usina = key
            end
        end
    end
end

vazao = Vazao.new(18, 'vazao.xls')
vazao.read_file