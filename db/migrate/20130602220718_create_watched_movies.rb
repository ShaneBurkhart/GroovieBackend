class CreateWatchedMovies < ActiveRecord::Migration
  def change
    create_table :watched_movies do |t|
      t.integer :user_id
      t.integer :flixster_id

      t.timestamps
    end
  end
end
