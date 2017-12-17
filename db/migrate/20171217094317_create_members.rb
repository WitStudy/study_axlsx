class CreateMembers < ActiveRecord::Migration[5.1]
  def change
    create_table :members do |t|
      t.string :code
      t.string :name
      t.integer :gender
      t.integer :age
      t.integer :team_id

      t.timestamps
    end
  end
end
