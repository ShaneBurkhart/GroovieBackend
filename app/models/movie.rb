class Movie < ActiveRecord::Base
	require 'open-uri'
	require 'json'
  # attr_accessible :title, :body
  #ex. http://api.rottentomatoes.com/api/public/v1.0/movies/770672122.json?apikey=[your_api_key]
  BASE_URL = "http://api.rottentomatoes.com/api/public/v1.0/"
  ROTTON_TOMATOES_API_KEY = "c2dhvn6k64escxu2s77bktkk"

  def self.get_movie(flixster_id)
  	JSON.parse(open(construct_url("movies/#{flixster_id}")).read)
  end

  def self.get_upcoming
  	JSON.parse(open(construct_url("lists/movies/upcoming")).read)
  end

  def self.get_in_theaters
  	JSON.parse(open(construct_url("lists/movies/in_theaters")).read)
  end

  private
	  def self.construct_url(api)
      #http://api.rottentomatoes.com/api/public/v1.0/lists/movies/upcoming.json?apikey=[your_api_key]
	  	"#{BASE_URL}#{api}.json?apikey=#{ROTTON_TOMATOES_API_KEY}"
		end  	
end
