require 'spec_helper'

describe "Users" do
  describe "create a user" do

    before do
     @user = User.new
     @user.name = "Shane"
     @user.email = "shaneburkhart@gmail.com"
     @user.password = "password"
   	end

   	subject { @user }

   	it { should respond_to :name }
   	it { should respond_to :email }
   	it { should respond_to :password }
   	it { should respond_to :has_role}
   	it { should respond_to :add_role}
   	it { should be_valid }

   	describe "password too short" do
   		before { @user.password = "a" }
   		it { should_not be_valid }
   	end

   	describe "password too long" do
   		before { @user.password = ('a' * 40) }
   		it { should_not be_valid }
   	end

   	describe "blank name" do
   		before { @user.name = " " }
   		it { should_not be_valid }
   	end

   	describe "blank email" do
   		before { @user.email = " " }
   		it { should_not be_valid }
   	end

   	describe "blank password" do
   		before { @user.password = " " }
   		it { should_not be_valid }
   	end

   	describe "when email format is invalid" do
	    it "should be invalid" do
	      addresses = %w[user@foo,com user_at_foo.org example.user@foo.
	                     foo@bar_baz.com foo@bar+baz.com]
	      addresses.each do |invalid_address|
	        @user.email = invalid_address
	        should_not be_valid
	      end      
	    end
	  end

	  describe "when email format is valid" do
	    it "should be valid" do
	      addresses = %w[user@foo.COM A_US-ER@f.b.org frst.lst@foo.jp a+b@baz.cn]
	      addresses.each do |valid_address|
	        @user.email = valid_address
	        should be_valid
	      end      
	    end
	  end

  end
end
