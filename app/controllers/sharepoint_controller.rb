require "addressable/uri"

class SharepointController < ApplicationController
  protect_from_forgery except: :auth

  def auth
    @spAppToken = params[:SPAppToken]
    @decoded = JWT.decode(@spAppToken, 'dnh0nrYHiLseGRzYID71fHGHbshr+lUUANbq1dKUlPQ=', false)

    @appctx = @decoded['appctx']
    @acsServer = JSON.parse(@appctx)['SecurityTokenServiceUri']

    part1 = '00000003-0000-0ff1-ce00-000000000000'
    part2 = 'fluxxtest.packard.org'
    part3 = @decoded['appctxsender'].split('@')[1]

    postdata = {'grant_type' => 'refresh_token',
      'client_id' => @decoded['aud'],
      'client_secret' => 'dnh0nrYHiLseGRzYID71fHGHbshr+lUUANbq1dKUlPQ=',
      'refresh_token' => @decoded['refreshtoken'],
      'resource' => part1 + '/' + part2 + '@' + part3}

    @result = RestClient.post @acsServer, postdata
    @access_token = JSON.parse(@result)['access_token']

    session[:sharepoint_access_token] = @access_token

    redirect_to sharepoint_url
  end

  def index
  end

  def create_folder
    resource = RestClient::Resource.new('https://fluxxtest.packard.org/_api/web/folders')
    response = resource.post(
      "{ '__metadata': { 'type': 'SP.Folder' }, 'ServerRelativeUrl': '/Proposal Documents/Dark Seid' }",
      'Authorization' => 'Bearer ' + session[:sharepoint_access_token],
      'X-RequestDigest' => x_request_digest,
      :accept => "application/json;odata=verbose",
      :content_type => "application/json;odata=verbose"
    )
    flash[:notice] = "Folder created successfully."

    redirect_to sharepoint_url
  rescue => e
    Rails.logger.error e.response
    raise e
  end

  def upload_file
    url = "https://fluxxtest.packard.org/_api/web/GetFolderByServerRelativeUrl('/Proposal Documents/Dark Seid')/Files/add(url='beo-dat-may-troi-piano.pdf',overwrite=true)"
    url = Addressable::URI.encode(url)

    request = RestClient::Request.new(
      method: :post,
      headers: {
        'Authorization' => 'Bearer ' + session[:sharepoint_access_token],
        'X-RequestDigest' => x_request_digest
      },
      url: url,
      payload: {
        multipart: true,
        file: File.new("/home/hoa/Downloads/beo-dat-may-troi-piano.pdf", 'rb')
      }
    )
    response = request.execute
    flash[:notice] = "File uploaded successfully."

    redirect_to sharepoint_url
  rescue => e
    # Rails.logger.error e.response
    raise e
  end

  def title
    Rails.logger.info "Call me !!!"
    auth = { 'Authorization' => 'Bearer ' + session[:sharepoint_access_token] }

    begin
      response = RestClient.get 'https://fluxxtest.packard.org/_api/web/title', auth
    rescue => e
      Rails.logger.error e.response.inspect
    end

    @title = Nokogiri.XML(response.body).xpath('//d:Title').first.text
  end

  private

  def x_request_digest
    @x_request_digest ||= begin
      resource = RestClient::Resource.new('https://fluxxtest.packard.org/_api/contextinfo')

      response = resource.post(
        '',
        'Authorization' => 'Bearer ' + session[:sharepoint_access_token],
        :accept => "application/json;odata=verbose"
      )

      parsed_body = JSON.parse(response.body)

      Rails.logger.info("X-RequestDigest: #{parsed_body['d']['GetContextWebInformation']['FormDigestValue']}")
      parsed_body['d']['GetContextWebInformation']['FormDigestValue']
    end
  rescue => e
    render :text => e.response
  end
end

