module ApplicationHelper
	BASE_TITLE = 'Movie App'

	def full_title(title)
		if !title.empty?
			"#{title} | #{BASE_TITLE}"
		else
			BASE_TITLE
		end
	end
end
