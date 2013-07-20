class WatchlistsController < ApplicationController
  def index
    @movies = WatchedMovies.find_by_user(current_user)
  end

  def create
  end

  def new
  end

  def update
  end

  def destroy
  end
end
