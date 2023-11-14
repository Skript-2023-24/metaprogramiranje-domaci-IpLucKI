require "google_drive"

$worksheet_1st = nil
$worksheet_2nd = nil
$worksheet_sum = nil
$worksheet_difference = nil

class String

    def is_integer?
        Integer(self)
        true
    rescue ArgumentError, TypeError
        false    
    end

    def is_float?
        Float(self)
        true
    rescue ArgumentError, TypeError
        false
    end
end

class Table
    include Enumerable

    attr_accessor :v_data, :hash_of_header

    def initialize
        @v_data = []
        @hash_of_header = {}        
    end

    def row(index)
        v_data[index - 1]
    end

    def each
        v_data.each do |arr|
            arr.each do |x|
                yield x
            end
        end
    end

    def [](column_name)
        v_column = []
        j = hash_of_header[column_name]        
        # j -= hash_of_header.first[1]
        v_data.each do |arr|            
            v_column << arr[j]
        end
        #v_column
        ColumnIndexingHelper.new(self, column_name, v_column)
    end    

    def +(t2)
        t1 = self
        # Provera da li su header-i isti        
        if t1.hash_of_header.keys != t2.hash_of_header.keys
            return nil
        end        
        res = Table.new
        res.hash_of_header = t1.hash_of_header
        t1.v_data.each do |arr|
            res.v_data << arr
        end
        t2.v_data.each do |arr|
            res.v_data << arr if !res.v_data.include?(arr)
        end        
        res
    end

    def -(t2)
        t1 = self
        if t1.hash_of_header.keys != t2.hash_of_header.keys
            return nil
        end
        res = Table.new
        res.hash_of_header = t1.hash_of_header
        t1.v_data.each do |arr|
            res.v_data << arr
        end        
        t2.v_data.each do |arr|
            res.v_data.reject! {|x| x == arr}
        end        
        res
    end

    class ColumnIndexingHelper
        attr_accessor :table, :column_name, :arr

        def initialize(table, column_name, arr)
            @table = table
            @column_name = column_name
            @arr = arr
        end

        def [](index)
            arr[index]
        end

        def []=(index, val)            
            j = table.hash_of_header[column_name]
            # j -= table.hash_of_header.first[1]            
            table.v_data[index][j] = val
        end        

        def to_s
            arr
        end
    end

    def retrieve_column_array(index)
        v_column = []
        v_data.each do |arr|
            v_column << arr[index]
        end
        v_column
    end

    def define_method_for_column
        hash_of_header.each do |key, value|
            method_name = "_" + key.downcase.gsub(/\s+/, "_")
            # Obican pristup odredjenoj koloni
            define_singleton_method(method_name) do
                retrieve_column_array(value)
            end
            # Izracunavanje necega za odredjenu kolonu
            define_singleton_method(method_name) do
                ColumnCalculatingHelper.new(self, retrieve_column_array(value))
            end
        end
    end

    class ColumnCalculatingHelper
        attr_accessor :table, :v_column

        def initialize(table, v_column)
            @v_column = v_column
            @table = table
        end

        def sum
            v_column.reduce(0) {|res, val| val.nil? ? res : res + val}
        end

        def avg
            (v_column.reduce(0) {|res, val| val.nil? ? res : res + val}).to_f / v_column.length
        end

        def method_missing(method_name)
            unless @method_missing_called
                @method_missing_called = true
                # Nadjem index row kroz column i onda samo uletim u v_data od glavne klase
                i = v_column.find_index(method_name.to_s)                
                result = i.nil? ? [] : table.v_data[i]
                @method_missing_called = false
                result
            end
        end        

        def map
            # Ignorisem nil polja
            result = []
            v_column.each do |x|
                !x.nil? ? result << (yield x) : result << x
            end
            result            
        end

        def select
            result = []
            v_column.each do |x|
                result << x if yield x
            end
            result
        end

        def reduce(init_val)
            result = init_val
            v_column.each do |x|
                result = yield(result, x)
            end
            result
        end

        def to_s
            v_column
        end
    end

    def to_s
        s = ""
        v_data.each do |arr|
            arr.each do |x|
                s += x.nil? ? "nil" : x.to_s
                s += " "
            end
            s += "\n"
        end
        s
    end
end

def find_start(t, worksheet)
    (1..worksheet.num_rows).each do |i|
        (1..worksheet.num_cols).each do |j|
            return [i, j] if !worksheet[i, j].empty?        
        end
    end
    [0, 0]
end

def extract_header(t, worksheet, i, j)
    val = 0
    while !worksheet[i, j].empty?
        # Necu da mi key bude redni broj kolone u samom sheet-u nego vezano za memoriju, odnosno array
        # A ok ipak moram da bih onda mogao da fizicki promenim cell u Google Sheets-u
        # Sto samo znaci da pri izracunavanju u samoj memoriji za array oduzmem od prvog elementa hash mape
        # A pa samo posto izgleda nismo vodili racuna o praznim redovima onda ne smemo da real time menjamo polja ddx
        # t.hash_of_header[$worksheet[i, j]] = j
        t.hash_of_header[worksheet[i, j]] = val
        val += 1
        j += 1
    end
end

def is_empty_row(worksheet, i, j, k)    
    counter = 0
    (j...j+k).each do |x|
        counter += 1 if worksheet[i, x].empty?            
        return true if counter == k
    end    
    false
end

def extract_data(t, worksheet, i_start, j_start)
    # Prvo moram da proverim da li je row prazan
    # Ukoliko nije prazan, a neka polja su prazna, njih cu tretirati kao 0 ili NaN to cu da vidim
    num_columns = t.hash_of_header.length
    (i_start..worksheet.num_rows).each do |i|
        if !is_empty_row(worksheet, i, j_start, num_columns) && !(worksheet[i, j_start - 1] =~ /total|subtotal/)
            arr = []
            (j_start...j_start+num_columns).each do |j|
                # Ovde pada jedan ddx convert momenat
                # arr << ($worksheet[i, j].empty? ? nil : $worksheet[i, j].to_i)
                # lule = nil
                if worksheet[i, j].empty?
                    lule = nil
                elsif worksheet[i, j].is_integer?
                    lule = worksheet[i, j].to_i
                elsif worksheet[i, j].is_float?
                    lule = worksheet[i, j].to_f
                else
                    lule = worksheet[i, j].to_s
                end
                arr << lule
            end
            t.v_data << arr
        end
    end    
end

def populate_table(t, worksheet)    
    indexes = find_start(t, worksheet)
    if indexes == [0, 0]
        p "Prazan je sheet"
        return
    end
    # p "Row: #{indexes[0]} Column: #{indexes[1]}"
    extract_header(t, worksheet, indexes[0], indexes[1])
    extract_data(t, worksheet, indexes[0] + 1, indexes[1])
end

def write_res(res, worksheet)
    return if res.nil?
    i = 2
    j_start = 2
    j = j_start
    # Prvo pisemo header
    res.hash_of_header.each do |key, value|
        worksheet[i, j] = key
        j += 1
    end
    i += 1
    j = j_start
    # Sada pisemo data
    res.v_data.each do |arr|
        arr.each do |x|
            worksheet[i, j] = x
            j += 1
        end
        i += 1
        j = j_start
    end
    worksheet.save
end

def main()    
    session = GoogleDrive::Session.from_config("config.json")
    $worksheet_1st = session.spreadsheet_by_key("1SjyE02CbblGj6rlHYuy01MbDYOg48JLCQueUdiec5wk").worksheets[0]
    $worksheet_2nd = session.spreadsheet_by_key("1SjyE02CbblGj6rlHYuy01MbDYOg48JLCQueUdiec5wk").worksheets[1]
    $worksheet_sum = session.spreadsheet_by_key("1SjyE02CbblGj6rlHYuy01MbDYOg48JLCQueUdiec5wk").worksheets[2]
    $worksheet_difference = session.spreadsheet_by_key("1SjyE02CbblGj6rlHYuy01MbDYOg48JLCQueUdiec5wk").worksheets[3]
    t1 = Table.new
    populate_table(t1, $worksheet_1st)
    puts "Funkcionalnost row:"
    p t1.row(3)
    puts "Funkcionalnost each:"
    t1.each do |x|
        p x
    end
    puts "Sintaksa t[\"Kolona\"]:"
    p t1["3rd"].to_s
    p t1["Druga Kolona"].to_s
    puts "Sintaksa t[\"Kolona\"][index]:"
    p t1["3rd"][1]
    p t1["Druga Kolona"][0]
    puts "Sintaksa dodele po gornjem pristupu celije:"
    p t1["3rd"][0]
    t1["3rd"][0] = 2000
    p t1["3rd"][0]
    puts "Sintaksa za direktan pristup kolonama:"
    t1.define_method_for_column
    p t1._index.to_s
    puts "Sintaksa za sum i avg nad gornjim pristupom:"
    p t1._druga_kolona.sum
    p t1._3rd.avg
    puts "Sintaksa za izvlacenje reda na osnovu vrednosti jedne od celija:"
    p t1._index.rn2031
    puts "Prikaz funkcionalnosti funkcija map, select, reduce:"
    p t1._3rd.map {|x| x += 2.5}
    p t1._3rd.select {|x| !x.nil? && x > 8}
    p t1._druga_kolona.reduce(10000) {|res, x| !x.nil? ? res - x : res}
    puts "Na osnovu samog generisanja tabele se moze videti da ignorisemo red ukoliko on ima total/subtotal!"
    puts "Prikaz operatora + nad tabelama:"
    t2 = Table.new
    populate_table(t2, $worksheet_2nd)
    res = t1 + t2
    puts res.nil? ? "Operacija je neuspesna jer tabele nemaju isti header" : res.to_s
    write_res(res, $worksheet_sum)
    puts "Prikaz operatora - nad tabelama:"
    res = t1 - t2
    puts res.nil? ? "Operacija je neuspesna jer tabele nemaju isti header" : res.to_s
    write_res(res, $worksheet_difference)
    puts "Na osnovu samog generisanja tabele se moze videti da ignorisemo prazne redove!"
end

main()


