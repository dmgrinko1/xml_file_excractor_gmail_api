class Gmail::XmlFileExtractor
  require 'google/apis/gmail_v1'
  require 'googleauth'
  require 'googleauth/stores/redis_token_store'
  require 'fileutils'

  TMP_DIR = "#{Rails.root}/tmp/redmart"
  OOB_URI = 'urn:ietf:wg:oauth:2.0:oob'.freeze
  APPLICATION_NAME = 'MyApp Mail Integration'.freeze
  GOOGLE_EMAIL_ADDRESS = ENV['GOOGLE_EMAIL_ADDRESS']

  Gmail = Google::Apis::GmailV1
  SYNCH_TYPE = 'gmail'

  SCOPE = Google::Apis::GmailV1::AUTH_GMAIL_MODIFY
  FROM_EMAIL_ADDRESS = ENV['FROM_EMAIL_ADDRESS']
  FROM_SUBJECT = ENV['FROM_SUBJECT']

  QUERY_FILTER = "from:#{FROM_EMAIL_ADDRESS} has:attachment subject:#{FROM_SUBJECT}"
  INBOX_LABEL_FILTER = %w[INBOX].freeze

  FAILED_LABEL = ENV['GOOGLE_FAILED_LABEL'].split('|')
  PROCESSED_LABEL = ENV['GOOGLE_PROCESSED_LABEL'].split('|')
  REMOVE_LABEL = %w[UNREAD IMPORTANT INBOX].freeze

  def initialize
    @google_client_id = ENV['GOOGLE_CLIENT_ID']
    @google_client_secret = ENV['GOOGLE_CLIENT_SECRET']
    FileUtils.rm_r(TMP_DIR) if File.directory?(TMP_DIR)
    FileUtils.mkdir_p(TMP_DIR)
  end

  def self.call
    new.call
  end

  def call
    MyApp.logger.info(I18n.t('api.gmail_uploader.no_messages')) && return if messages.nil?
    channel = SalesChannel.find_by_channel_id('channel_id')
    spreadsheet = Roo::Spreadsheet.open(upload_file_path)
    header = spreadsheet.row(1)
    prefix_separator = '-'
    if spreadsheet.last_row <= 1
      spreadsheet.close
      MyApp.logger.info(I18n.t('api.gmail_uploader.no_orders'))
      return
    end
    modify_label(add_label_ids: FAILED_LABEL, remove_label_ids: REMOVE_LABEL)
    attachment_is_valid =
      ::Gmail::AttachmentChecker.new(channel,
                                     spreadsheet,
                                     header,
                                     prefix_separator,
                                     email_date).call
    unless attachment_is_valid == true
      raise StandardError, I18n.t('api.gmail_uploader.errors.something_went_wrong', date: email_date)
    end

    ::Order::ExternalSalesOrderService.new.all(
      upload_file_path,
      channel,
      prefix_separator,
      synch_type: SYNCH_TYPE
    )

    after_success
  end

  private

  def find_file
    "#{TMP_DIR}/redmart_report.xlsx"
  end

  def upload_file_path
    find_file.tap do |path|
      File.open(path, 'wb') { |f| f.puts attachment.data }
    end
  end

  def modify_label(add_label_ids:, remove_label_ids:)
    option = Google::Apis::GmailV1::ModifyMessageRequest.new(
      add_label_ids: add_label_ids,
      remove_label_ids: remove_label_ids
    )
    gmail.modify_message(GOOGLE_EMAIL_ADDRESS, message_id, option)
  end

  def attachment
    gmail.get_user_message_attachment(GOOGLE_EMAIL_ADDRESS, message_id, attachment_id)
  end

  def label
    gmail.list_user_labels(GOOGLE_EMAIL_ADDRESS)
  end

  def attachment_id
    filtered_attachment = email.payload.parts.
                          find { |p| p.filename.include?('.xlsx') && p.filename.include?('Pick Up Request') }
    if filtered_attachment.present?
      filtered_attachment.body.attachment_id
    else
      modify_label(add_label_ids: FAILED_LABEL, remove_label_ids: REMOVE_LABEL)
      raise StandardError, I18n.t('api.gmail_uploader.errors.bad_attachment')
    end
  end

  def messages
    gmail.list_user_messages(GOOGLE_EMAIL_ADDRESS, q: QUERY_FILTER, label_ids: INBOX_LABEL_FILTER).messages
  end

  def message_id
    @message_id ||= messages.first.id
  end

  def email
    gmail.get_user_message(GOOGLE_EMAIL_ADDRESS, message_id)
  end

  def email_date
    email.payload.headers[1].value
  end

  def gmail
    @gmail ||= Gmail::GmailService.new.tap do |publisher|
      publisher.client_options.application_name = APPLICATION_NAME
      publisher.authorization = authorize
    end
  end

  def authorize
    client_id = Google::Auth::ClientId.new(@google_client_id, @google_client_secret)
    token_store = Google::Auth::Stores::RedisTokenStore.new
    authorizer = Google::Auth::UserAuthorizer.new client_id, SCOPE, token_store
    user_id = 'redmart_pick_up_request'

    credentials = authorizer.get_credentials user_id
    # credentials = nil

    if credentials.nil?
      url = authorizer.get_authorization_url base_url: OOB_URI
      puts 'Open the following URL in the browser and enter the ' \
           'resulting code after authorization:\n' + url
      code = gets
      credentials = authorizer.get_and_store_credentials_from_code(
        user_id: user_id, code: code, base_url: OOB_URI
      )
    end
    credentials
  end

  def after_success
    File.delete(find_file)
    delete_labels = (REMOVE_LABEL + FAILED_LABEL)
    modify_label(add_label_ids: PROCESSED_LABEL, remove_label_ids: delete_labels)
    MyApp.logger.info(I18n.t('api.gmail_uploader.success.uploaded',
                                 date: email_date))
  end
end
