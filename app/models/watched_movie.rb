class WatchedMovie < ActiveRecord::Base
	belongs_to :user
  attr_accessible :flixster_id, :user_id
  validates :flixster_id, :user_id, presence: true, uniqueness: true,
  				numericality: { greater_than: 0 }
end
