require 'spec_helper'

describe "WatchedMovies" do
	describe "WatchedMovie model test" do

		before { @movie = WatchedMovie.new }

		subject { @movie }

		it { should respond_to :user }
		it { should respond_to :flixster_id }

		describe "blank user_id" do
			before { @movie.user_id = nil }
			it { should_not be_valid }
		end

		describe "blank flixster_id" do
			before { @movie.flixster_id = nil }
			it { should_not be_valid}
		end

		describe "zero user_id" do
			before { @movie.user_id = 0 }
			it { should_not be_valid }
		end

		describe "zero flixster_id" do
			before { @movie.flixster_id = 0 }
			it { should_not be_valid}
		end

		describe "negative user_id" do
			before { @movie.user_id = -1 }
			it { should_not be_valid }
		end

		describe "negative flixster_id" do
			before { @movie.flixster_id = -1 }
			it { should_not be_valid}
		end

		describe "valid check" do
			before do
				@movie.user_id = 1
				@movie.flixster_id = 1
			end
			it { should be_valid }
		end

		describe "no duplicates" do
			before do
				@movie.save
				dup = @movie.dup 
			end
			it { should_not be_valid }
		end

		describe "build relationship with user" do
			before do
				@user = User.new
				@user.name = "Shane"
				@user.email = "user@gmail.com"
				@user.password = "password"
				@user.save
			end
			it { @user.watched_movies.build.should_not == nil }
		end

	end
end
