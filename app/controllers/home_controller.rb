class HomeController < ApplicationController
  def index
  	@upcoming_theater = Movie.get_upcoming
  	#@in_theater = Movie.get_in_theaters
  end
end
