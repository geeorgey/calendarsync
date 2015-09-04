#!/usr/bin/ruby

require 'rubygems'
require 'active_support'
require 'active_support/core_ext'
require 'google/api_client'
require 'google/api_client/client_secrets'
require 'google/api_client/auth/installed_app'
require 'google/api_client/auth/storage'
require 'google/api_client/auth/storages/file_store'
require 'fileutils'
require 'csv'
require 'open3'


class NilClass
  def method_missing(name, *args, &block)
      nil
        end
        end

APPLICATION_NAME = 'Google Calendar API Quickstart'
CLIENT_SECRETS_PATH = '/Users/george/Downloads/client_secret.json'
CREDENTIALS_PATH = File.join(Dir.home, '.credentials',
                             "calendar-quickstart.json")
                             SCOPE = 'https://www.googleapis.com/auth/calendar.readonly'

##
# Ensure valid credentials, either by restoring from the saved credentials
# files or intitiating an OAuth2 authorization request via InstalledAppFlow.
# If authorization is required, the user's default browser will be launched
# to approve the request.
#
# @return [Signet::OAuth2::Client] OAuth2 credentials
def authorize
  FileUtils.mkdir_p(File.dirname(CREDENTIALS_PATH))

  file_store = Google::APIClient::FileStore.new(CREDENTIALS_PATH)
    storage = Google::APIClient::Storage.new(file_store)
      auth = storage.authorize

  if auth.nil? || (auth.expired? && auth.refresh_token.nil?)
      app_info = Google::APIClient::ClientSecrets.load(CLIENT_SECRETS_PATH)
          flow = Google::APIClient::InstalledAppFlow.new({
                :client_id => app_info.client_id,
                      :client_secret => app_info.client_secret,
                            :scope => SCOPE})
                                auth = flow.authorize(storage)
                                    puts "Credentials saved to #{CREDENTIALS_PATH}" unless auth.nil?
                                      end
                                        auth
                                        end

userList = CSV.read("/Users/yourname/googleAppsUserList.csv") #メアドのリストです

# Initialize the API
client = Google::APIClient.new(:application_name => APPLICATION_NAME)
client.authorization = authorize
calendar_api = client.discovered_api('calendar', 'v3')

t = Time.now - 3.minutes #3分毎に更新するのでこの設定にした
t_max = Time.now + 3.month #三ヶ月先までにしましたが、おこのみで


File.open("/Users/yourname/output.csv","w") do |file| #output.csvは毎回空にしてからデータ入れる
  file.puts '"ID","Email","googleCalEventID__c","Deleted","StartDateTime","EndDateTime","Subject","Description","Location"'
  end

  for uid in userList do
      # Fetch the next 10 events for the user
          results = client.execute!(
                :api_method => calendar_api.events.list,
                      :parameters => {
                              :calendarId => uid,
                                      :maxResults => 20,
                                              :orderBy => 'updated',
                                                      :updatedMin => t.iso8601 ,
                                                              :timeMin => Time.now.iso8601,
                                                                      :timeMax => t_max.iso8601
                                                                              }
                                                                                      )

 
     File.open("/Users/yourname/output.csv","a") do |file|
         results.data.items.each do |event|
               start = (event.start.date || event.start.date_time) - 9.hours #GMT +9の解消
                     end_time = (event.end.date || event.end.date_time) - 9.hours #GMT +9の解消
                           start2 = start.iso8601
                           #ここの処理ですが、Googleカレンダー上のデータを削除した時に、Salesforceのデータも削除したかったのですが、ISDELETEDフラグはスクリプトからいじれないようです。その為、削除されたデータははるか昔の日付に飛ばす処理にしてあります
                                 if start.nil? then
                                         start2 = "2000-08-03T08:00:00.000Z"
                                               else
                                                       start2 = start.iso8601
                                                             end
                                                                   if end_time.nil? then
                                                                           end_time2 = "2000-08-04T08:00:00.000Z"
                                                                                 else
                                                                                         end_time2 = end_time.iso8601
                                                                                               end
                                                                                               #ここまでが削除時の処理
                                                                                               #ここの処理ブサイクなので誰か教えてください…ダブルクオーテーションで囲むスマートな方法が知りたい
                                                                                                     file.print '""',",#{uid},#{uid}_#{event.id}",'","',"#{event.status}",'","',"#{start2}",'","',"#{end_time2}" ,'","',"#{event.summary}",'","',"#{event.description}",'","',"#{event.location}",'"',"\n"
                                                                                                         end

    end
      end

  #置換。余計な文字の削除と、Salesforceに取り込む際の時間データの型を整形しています。
    f = File.open("/Users/yourname/output.csv","r")
      buffer = f.read();
        #buffer.gsub!(/"|[|]| +0900| UTC/,'"' =>'','[' => '',']' => '','+09:00' => '.000Z',' UTC' => '')
          #buffer.gsub!('"',"").gsub!('[',"").gsub!(']',"").gsub!('+09:00',".000Z").gsub!(' UTC',"");
            buffer.gsub!('[',"").gsub!(']',"").gsub!('lne.st"_',"lne.st_").gsub!('+09:00',".000Z").gsub!(' UTC',"").gsub!('confirmed',"FALSE").gsub!('cancelled',"TRUE");
              f=File.open("/Users/george/GoogleDrive/forTalend/googleCal/output.csv","w")
                f.write(buffer)
                  f.close()


  File.readlines("/Users/yourname/compare.csv").uniq
    compare = CSV.table('/Users/yourname/compare.csv', force_quotes: true)
      output = CSV.table('/Users/yourname/output.csv', force_quotes: true)

  # Talendに喰わせる更新用ファイルの初期化
    File.open("/Users/yourname/updatedEvents.csv","w") do |file|
        file.puts '"ID","Email","googleCalEventID__c","Deleted","StartDateTime","EndDateTime","Subject","Description","Location"'
          end
            File.open("/Users/yourname/noID.csv","w") do |file|
                file.puts '"ID","Email","googleCalEventID__c","Deleted","StartDateTime","EndDateTime","Subject","Description","Location"'
                  end

  # ファイルへのデータ書き込み
    output.each{|googlecaleventid__c|
        if compare[:googlecaleventid__c].include?(googlecaleventid__c[2]) then
              #存在データを更新するための出力
                    index = compare.find_index{|x| #既存データのindexを取得してSFのIDを取ってくる処理
                          x[:googlecaleventid__c] == googlecaleventid__c[2]
                                }

      File.open("/Users/yourname/updatedEvents.csv","a") do |file|
            #  file.puts "ID,Email,googleCalEventID__c,Subject,StartDateTime,EndDateTime,googleCalEventLastUpdated__c,Description,Location"
            #この出力処理もブサイクなので何とかしたい。SFのIDとGoogleカレンダーの更新情報を組み合わせています
                    file.print '"', compare[index][:id] ,'","',  googlecaleventid__c[1] ,'","', googlecaleventid__c[2] ,'","', googlecaleventid__c[3] ,'","', googlecaleventid__c[4] ,'","', googlecaleventid__c[5]  ,'","',  googlecaleventid__c[6],'","',  googlecaleventid__c[7],'","',  googlecaleventid__c[8],'","', googlecaleventid__c[9] , '"' , "\n"
                          end

    else
          #新規データを出力する
                File.open("/Users/yourname/noID.csv","a") do |file|
                      #  file.puts "ID,Email,googleCalEventID__c,Subject,StartDateTime,EndDateTime,googleCalEventLastUpdated__c,Description,Location"
                      #ここもどうにかしたい
                              file.print '"","' ,googlecaleventid__c[1] ,'","', googlecaleventid__c[2] ,'","', googlecaleventid__c[3] ,'","', googlecaleventid__c[4] ,'","', googlecaleventid__c[5]  ,'","',  googlecaleventid__c[6],'","',  googlecaleventid__c[7],'","',  googlecaleventid__c[8], '"' , "\n"
                                    end
                                        end
                                          }


#RubyからTalendの出力ジョブを実行する処理

  DIR = "/Users/yourname/"

  cmd = 'cd ' + DIR + ";"
    cmd += 'ROOT_PATH=' + DIR + ";"
      cmd += 'java -Xms256M -Xmx1024M -cp $ROOT_PATH:$ROOT_PATH/../lib/systemRoutines.jar::$ROOT_PATH/../lib/userRoutines.jar::.:$ROOT_PATH/googlecalupdates_0_1.jar:$ROOT_PATH/../lib/activation-1.1.jar:$ROOT_PATH/../lib/advancedPersistentLookupLib-1.0.jar:$ROOT_PATH/../lib/axiom-api-1.2.13.jar:$ROOT_PATH/../lib/axiom-impl-1.2.13.jar:$ROOT_PATH/../lib/axis2-adb-1.6.2.jar:$ROOT_PATH/../lib/axis2-kernel-1.6.2.jar:$ROOT_PATH/../lib/axis2-transport-http-1.6.2.jar:$ROOT_PATH/../lib/axis2-transport-local-1.6.2.jar:$ROOT_PATH/../lib/commons-codec-1.3.jar:$ROOT_PATH/../lib/commons-collections-3.2.jar:$ROOT_PATH/../lib/commons-httpclient-3.1.jar:$ROOT_PATH/../lib/commons-logging-1.1.1.jar:$ROOT_PATH/../lib/dom4j-1.6.1.jar:$ROOT_PATH/../lib/geronimo-stax-api_1.0_spec-1.0.1.jar:$ROOT_PATH/../lib/httpcore-4.2.1.jar:$ROOT_PATH/../lib/jboss-serialization.jar:$ROOT_PATH/../lib/log4j-1.2.15.jar:$ROOT_PATH/../lib/mail-1.4.jar:$ROOT_PATH/../lib/neethi-3.0.1.jar:$ROOT_PATH/../lib/salesforceCRMManagement.jar:$ROOT_PATH/../lib/talend_file_enhanced_20070724.jar:$ROOT_PATH/../lib/talendcsv.jar:$ROOT_PATH/../lib/trove.jar:$ROOT_PATH/../lib/wsdl4j-1.6.3.jar:$ROOT_PATH/../lib/wstx-asl-3.2.9.jar:$ROOT_PATH/../lib/xmlschema-core-2.0.1.jar: salesforce.googlecalupdates_0_1.googleCalUpdates --context=Default "$@" '

