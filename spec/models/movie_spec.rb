require 'spec_helper'

describe Movie do
  describe "Test for methods" do
  	it{ Movie.should respond_to(:get_movie) }
  	it{ Movie.should respond_to(:get_upcoming) }
  	it{ Movie.should respond_to(:get_in_theaters) }
  	it{ Movie.should respond_to(:top_rentals) }
  	it{ Movie.should respond_to(:get_current_dvds) }
  	it{ Movie.should respond_to(:get_upcoming_dvds) }
  	it{ Movie.should respond_to(:get_new_release) }
  end

  describe "Test to make sure not nil" do
	end
end
