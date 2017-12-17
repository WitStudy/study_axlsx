# This file should contain all the record creation needed to seed the database with its default values.
# The data can then be loaded with the rails db:seed command (or created alongside the database with db:setup).
#
# Examples:
#
#   movies = Movie.create([{ name: 'Star Wars' }, { name: 'Lord of the Rings' }])
#   Character.create(name: 'Luke', movie: movies.first)

# Team
Team.delete_all
('A'..'C').each do |alpha|
  Team.create(name: "#{alpha}チーム")
end

# Member
Member.delete_all
teams = Team.all
MEMBER_SIZE = teams.count * 10
(1..MEMBER_SIZE).each do |m|
  Member.create(
    code: "M#{format('%04d', m)}",
    name: "メンバー#{m}",
    gender: m % 2,
    age: 20 + rand(40),
    team_id: teams[m % teams.count].id
  )
end
