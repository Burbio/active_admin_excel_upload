module ActiveAdminExcelUpload
  module ExcelParsable
    extend ActiveSupport::Concern

    class_methods do
      def publish_to_channel(channel_name,message)
        message = "[#{self.to_s}]#{message}"
        ActionCable.server.broadcast channel_name, message
      end

      def excel_create_record(row, index, header,channel_name)
        self.publish_to_channel(channel_name, "processing for #{row}")
        object = Hash[header.zip row]
        record = self.new(object)
        if record.save
          self.publish_to_channel(channel_name,"Successfully created record for #{row}, id: #{record.id}")
        else
          # self.publish_to_channel(channel_name,"Could not create record for #{row}, error: #{record.errors.messages}")
          self.publish_to_channel(channel_name,more_delightful_message(row,record.errors.messages))
        end
      end

      def more_delightful_message(row, messages)
        first_message = row.kind_of?(String) ? row : row.try(:[],0)
        error_messages = messages[:error].map{|k,v| "<p>#{k.to_s} #{v.to_s}</p>"}.join("").html_safe
        "<p>Could not create record for row #{first_message}.</p>".html_safe + error_messages
      end


      def excel_process_sheet(sheet,current_admin_user)
        xlsx = Roo::Spreadsheet.open(sheet)
        sheet = xlsx.sheet(xlsx.sheets.index(self.table_name))
        header = sheet.row(1)
        channel_name = "excel_channel_#{current_admin_user.id}"
        logger.info "Brian :: #{channel_name}"
        header_downcase = header.map(&:parameterize).map(&:underscore)
        log_table_name(self)
        self.publish_to_channel(channel_name,"Start processing sheet #{self.table_name}")
        self.publish_to_channel(channel_name,"Start processing sheet #{header}")
        sheet.parse.each_with_index do |row, index|
          begin
            self.excel_create_record(row,index,header_downcase,channel_name)
          rescue StandardError => e
            self.publish_to_channel(channel_name,"Exception while processing #{row}, Exception: #{e.message}")
          end
        end
        self.publish_to_channel(channel_name, "End processing sheet #{self.table_name}")
      end

      def table_header
        %Q[<thead class="border-b">
            <tr>
              <th scope="col" class="text-sm font-medium text-gray-900 px-6 py-4 text-left">
                #
              </th>
              <th scope="col" class="text-sm font-medium text-gray-900 px-6 py-4 text-left">
                Heading
              </th>
              <th scope="col" class="text-sm font-medium text-gray-900 px-6 py-4 text-left">
                Heading
              </th>
              <th scope="col" class="text-sm font-medium text-gray-900 px-6 py-4 text-left">
                Heading
              </th>
              <th scope="col" class="text-sm font-medium text-gray-900 px-6 py-4 text-left">
                Heading
              </th>
              <th scope="col" class="text-sm font-medium text-gray-900 px-6 py-4 text-left">
                Heading
              </th>
              <th scope="col" class="text-sm font-medium text-gray-900 px-6 py-4 text-left">
                Heading
              </th>
              <th scope="col" class="text-sm font-medium text-gray-900 px-6 py-4 text-left">
                Heading
              </th>
              <th scope="col" class="text-sm font-medium text-gray-900 px-6 py-4 text-left">
                Heading
              </th>
            </tr>
          </thead>]
      end

      def log_table_name(itself)
        begin
          logger.info "#{itself.table_name}"
        rescue => e 
          logger.info "ANNABELLE -- JOB FAILED"
          logger.info itself 
          logger.info e.message
        end
      end
    end
  end
end
